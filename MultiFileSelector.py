import tkinter as tk
from tkinter import filedialog


class MultiFileSelector:
    """
    A GUI class for selecting multiple files and executing callbacks for analysis.

    """

    def __init__(self, master, test_callback, error_callback, max_files=20):
        """
        Initialize the MultiFileSelector class with GUI components.

        """
        self.master = master
        self.frame = tk.Frame(self.master)
        self.frame.grid(row=0, column=0)
        self.max_files = max_files
        self.filepaths = []

        self.test_callback = test_callback
        self.error_callback = error_callback

        self.label = tk.Label(self.frame, text="Select up to 20 files:")
        self.label.grid(row=0, column=0)

        self.select_button = tk.Button(self.frame, text="Select Files", command=self.select_files)
        self.select_button.grid(row=0, column=1)

        self.file_list_label = tk.Label(
            self.frame, text="", anchor="w", justify="left", wraplength=400
        )
        self.file_list_label.grid(row=1, column=0, columnspan=2, sticky="w")

        self.test_button = tk.Button(
            self.frame, text="Test Statistics", command=self.run_test_analysis, state="disabled"
        )
        self.test_button.grid(row=2, column=0)

        self.error_button = tk.Button(
            self.frame, text="Error Statistics", command=self.run_error_analysis, state="disabled"
        )
        self.error_button.grid(row=2, column=1)

    def select_files(self):
        """
        Open a file dialog to select multiple files and update the file list.
        """
        selected_files = filedialog.askopenfilenames(filetypes=[("HTML Files", "*.html")])

        for filepath in selected_files:
            if len(self.filepaths) < self.max_files:
                self.filepaths.append(filepath)
            else:
                break

        if self.filepaths:
            self.test_button["state"] = "normal"
            self.error_button["state"] = "normal"

        self.file_list_label.config(text="\n".join(self.filepaths))

    def run_test_analysis(self):
        """
        Run the test statistics analysis using the selected files.
        """
        self.test_callback(self.filepaths)

    def run_error_analysis(self):
        """
        Run the error statistics analysis using the selected files.
        """
        self.error_callback(self.filepaths)
