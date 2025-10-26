import pandas as pd
from collections import defaultdict
import os
import re
from Utils import add_campaign_details_rows, extract_campaign_details, parse_html, parse_issues


def extract_messages_from_files(filepaths):
    """
    Extract messages from multiple files and collect error statistics.

    """
    error_failure_data = []
    campaign_details = {}
    dates = []

    test_case_pattern = re.compile(r"^\d{2}_\d{2}$")

    for filepath in filepaths:
        soup = parse_html(filepath)

        campaign_date, details = extract_campaign_details(filepath, soup)
        campaign_details[filepath] = details
        dates.append(campaign_date)

        stimulations = []
        for tag in soup.find_all("span", class_=True):
            parsed_issue = parse_issues(tag, stimulations)
            if not parsed_issue:
                continue

            test_case_name = parsed_issue["test_case"]
            if test_case_pattern.match(test_case_name):
                continue

            error_failure_data.append({
                "Test Case": test_case_name,
                "Message": parsed_issue["message"].split(":", 1)[-1].strip(),
                "Category": parsed_issue["type"],
                "Date": campaign_date,
            })

    return error_failure_data, sorted(set(dates)), campaign_details


def prepare_error_failure_analysis(error_failure_data, dates):
    """
    Prepare error failure analysis data for reporting.

    """
    error_analysis = defaultdict(
        lambda: {"Occurrences": 0, "Test Cases": set(), "Date Counts": defaultdict(int)}
    )

    for entry in error_failure_data:
        message = entry["Message"]
        category = entry["Category"]
        test_case = entry["Test Case"]
        date = entry["Date"]

        error_analysis[message]["Occurrences"] += 1
        error_analysis[message]["Test Cases"].add(test_case)
        error_analysis[message]["Date Counts"][date] += 1
        error_analysis[message]["Category"] = category

    error_analysis_data = []
    for message, details in error_analysis.items():
        row = {
            "Error/Failure Message": message,
            "Category": details["Category"],
            "Occurrences": details["Occurrences"],
            "Associated Test Cases": "; ".join(sorted(details["Test Cases"])),
        }
        for date in dates:
            row[date] = details["Date Counts"].get(date, "--")
        error_analysis_data.append(row)

    return pd.DataFrame(error_analysis_data)


def generate_error_statistics(filepaths, save_path):
    """
    Generate a summary report for error statistics across multiple files.

    """
    error_failure_data, dates, campaign_details = extract_messages_from_files(filepaths)

    error_failure_df = prepare_error_failure_analysis(error_failure_data, dates)
    add_campaign_details_rows(error_failure_df, campaign_details, dates)
    output_file = os.path.join(save_path, "ErrorStatistics_Summary.xlsx")
    with pd.ExcelWriter(output_file, engine="openpyxl") as writer:
        error_failure_df.to_excel(writer, sheet_name="Error-Failure Analysis", index=False)

    print(f"Error statistics saved to {output_file}")
