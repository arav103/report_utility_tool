import pandas as pd
from collections import defaultdict
import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
from openpyxl.styles import Font
from Utils import (
    parse_html,
    parse_issues,
    process_nonduplicate
)

def extract_messages(html_file):
    """
    Extract messages from the provided HTML file.

    """
    soup = parse_html(html_file)

    passes = defaultdict(lambda: {"stimulations": [], "test_cases": []})
    issues = defaultdict(
        lambda: {
            "stimulations": [],
            "test_cases": [],
            "messages": [],
            "types": [],
            "times": [],
            "previous_actions": []
        }
    )
    warnings = defaultdict(lambda: {"stimulations": [], "test_cases": []})

    stimulations = [
        (div, div.find("span", class_="highlight").get_text(strip=True))
        for div in soup.find_all("div") if "title" in div.get("class", []) and div.find("span", class_="highlight")
    ]
    test_cases = [
        (div, div.get_text(strip=True).split()[0])
        for div in soup.find_all("div") if "title" in div.get("class", []) and "test" in div.get("class", [])
    ]

    for tag in soup.find_all("span", class_=True):
        issue = parse_issues(tag, stimulations)
        if issue and issue["test_case"] not in issues[issue["stimulation"]]["test_cases"]:
            issues[issue["stimulation"]]["stimulations"].append(issue["stimulation"])
            issues[issue["stimulation"]]["test_cases"].append(issue["test_case"])
            issues[issue["stimulation"]]["messages"].append(issue["message"])
            issues[issue["stimulation"]]["types"].append(issue["type"])
            issues[issue["stimulation"]]["times"].append(issue["timestamp"])
            issues[issue["stimulation"]]["previous_actions"].append(issue["previous_actions"])

    process_nonduplicate(soup, stimulations, test_cases, "PASS", passes)
    process_nonduplicate(soup, stimulations, test_cases, "WARNING", warnings)

    campaign_details = {}
    campaign_section = soup.find("div", {"data-tab": "campaign"})
    if campaign_section:
        for row in campaign_section.find_all("tr"):
            columns = row.find_all("td")
            if len(columns) == 2:
                key = columns[0].get_text(strip=True)
                value = columns[1].get_text(strip=True)
                campaign_details[key] = value

    return passes, issues, warnings, campaign_details

def clean_message(message):
    """
    Clean the message by extracting the relevant part after ':'.

    """
    if ':' in message:
        return message.split(':', 1)[1].strip()
    return message

def prepare_message_rows(data, is_pass=False):
    """
    Prepare rows for passes, warnings, or issues.

    """
    rows = []
    last_test_case = None
    for stim, info in data.items():
        for i in range(len(info["test_cases"])):
            test_case = info["test_cases"][i]
            current_test_case = "" if test_case == last_test_case else test_case
            if is_pass:
                rows.append([current_test_case, stim, ""])
            else:
                type_value = info["types"][i]
                message = clean_message(info["messages"][i]) if "messages" in info else ""
                time = info["times"][i]
                previous_actions = info["previous_actions"][i]
                rows.append([current_test_case, type_value, stim, message, time, previous_actions])
            last_test_case = test_case
    return rows

def generate_excel_report(html_file, passes, warnings, issues, campaign_details):
    """
     Generate an Excel report with passes, warnings, issues, and campaign details.

     """
    issues_data_frame = pd.DataFrame(
        prepare_message_rows(issues),
        columns=["Test Case", "Type", "Stimulation", "Message", "Time", "Previous Actions"]
    )
    issues_data_frame.insert(5, "Jira Ticket", "")
    issues_data_frame.insert(6, "Jira Status", "")
    issues_data_frame.insert(7, "Comments", "")

    passes_data_frame = pd.DataFrame(
        prepare_message_rows(passes, is_pass=True),
        columns=["Test Case", "Stimulation", "Message"]
    )

    warnings_data_frame = pd.DataFrame(
        prepare_message_rows(warnings, is_pass=True),
        columns=["Test Case", "Stimulation", "Message"]
    )
    pattern = r"^\d{2}_\d{2}$"
    issues_data_frame = issues_data_frame[~issues_data_frame["Test Case"].str.match(pattern, na=False)]

    if not passes_data_frame.empty:
        passes_data_frame = passes_data_frame.iloc[1:]

    if not warnings_data_frame.empty:
        warnings_data_frame = warnings_data_frame.iloc[1:]

    output_file = html_file + "_summary.xlsx"
    with pd.ExcelWriter(output_file, engine='openpyxl') as writer:

        issues_data_frame.to_excel(writer, sheet_name="Issues", index=False)
        passes_data_frame.to_excel(writer, sheet_name="Passes", index=False)
        warnings_data_frame.to_excel(writer, sheet_name="Warnings", index=False)

        category_counts = {
            "Passes": len(passes_data_frame),
            "Warnings": len(warnings_data_frame),
            "Failures": len(issues_data_frame[issues_data_frame["Type"] == "Failure"]),
            "Errors": len(issues_data_frame[issues_data_frame["Type"] == "Error"]),
        }

        fixed_categories = ["Passes", "Warnings", "Failures", "Errors"]
        fixed_colors = {"Passes": "green", "Warnings": "yellow", "Failures": "orange", "Errors": "red"}
        category_counts = {key: category_counts.get(key, 0) for key in fixed_categories}

        sorted_counts = [(key, category_counts[key]) for key in fixed_categories]
        categories = [x[0] for x in sorted_counts]
        counts = [x[1] for x in sorted_counts]
        colors = [fixed_colors[category] for category in categories]

        summary_data = pd.DataFrame(sorted_counts, columns=["Category", "Count"])
        summary_data.to_excel(writer, sheet_name="Summary", index=False, startrow=6)

        workbook = writer.book
        summary_sheet = workbook["Summary"]

        campaign_info = [
            ("Campaign Name", "Campaign name"),
            ("Campaign Date", "Campaign date"),
            ("Duration", "Duration"),
            ("ENNA Version", "ENNA version"),
            ("Python Version", "Python version"),
        ]

        for row, (label, key) in enumerate(campaign_info, start=1):
            summary_sheet.cell(row=row, column=1, value=label).font = Font(bold=True)
            summary_sheet.cell(row=row, column=2, value=campaign_details.get(key, "N/A"))

        plt.figure(figsize=(5, 5))
        plt.pie(counts, labels=categories, colors=colors, autopct='%1.1f%%', startangle=140)
        plt.title("Message Summary")
        plt.savefig("pie_chart.png")
        plt.close()

        img = Image("pie_chart.png")
        summary_sheet.add_image(img, "E7")

    print(f"Excel report with filtered data and pie chart written to {output_file}")

def analyze(html_file, save_path):
    """
    Analyze the HTML file and generate a report.

    """
    passes, issues, warnings, campaign_details = extract_messages(html_file)
    output_file = f"{save_path}/{html_file.split('/')[-1]}_summary.xlsx"
    generate_excel_report(output_file, passes, warnings, issues, campaign_details)