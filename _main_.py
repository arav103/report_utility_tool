import tkinter as tk
from tkinter import filedialog, messagebox
from SingleFileSelector import SingleFileSelector
from MultiFileSelector import MultiFileSelector
from SingleDayAnalysis import analyze
from CyclicRunAnalysis import analyze_cyclic_run
from MultipleFileAnalysis import generate_multi_file_summary
from ErrorStatistics import generate_error_statistics


def analyse_single(filepath, savepath, cyclic_run=False):
    """
    Analyze a single file and generate a report.

    """
    if cyclic_run:
        analyze_cyclic_run(filepath, savepath)
    else:
        analyze(filepath, savepath)

    messagebox.showinfo(
        "Report Generated",
        f"The report for the file '{filepath}' has been successfully saved to '{savepath}'."
    )


def analyse_test_statistics(filepaths):
    """
    Analyze multiple files for test statistics and generate a summary report.

    """
    savepath = filedialog.askdirectory(title="Select Folder to Save Test Statistics Report")
    if savepath:
        generate_multi_file_summary(filepaths, savepath)
        messagebox.showinfo(
            "Reports Generated",
            f"The test statistics report has been successfully saved to '{savepath}'."
        )


def analyse_error_statistics(filepaths):
    """
    Analyze multiple files for error statistics and generate a summary report.

    """
    savepath = filedialog.askdirectory(title="Select Folder to Save Error Statistics Report")
    if savepath:
        generate_error_statistics(filepaths, savepath)
        messagebox.showinfo(
            "Reports Generated",
            f"The error statistics report has been successfully saved to '{savepath}'."
        )


def select_mode():
    """
    Prompt the user to select the mode of operation: single file or multiple files.

    """
    mode = messagebox.askquestion(
        "Select Mode",
        "Would you like to process a single day file?"
    )

    if mode == 'yes':
        global fileselector
        fileselector = SingleFileSelector(
            root,
            run_function=analyse_single,
            cyclic_run_function=lambda filepath, savepath: analyse_single(filepath, savepath, cyclic_run=True)
        )
    else:
        global multifileselector
        multifileselector = MultiFileSelector(
            root,
            test_callback=analyse_test_statistics,
            error_callback=analyse_error_statistics
        )


if __name__ == '__main__':
    root = tk.Tk()
    root.minsize(400, 300)
    root.title("File Analysis Tool")

    select_mode()

    root.mainloop()
