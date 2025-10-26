import tkinter as tk
from tkinter import filedialog, messagebox


class SingleFileSelector:
    """
    A GUI class for selecting a single file for analysis and specifying a folder to save the report.

    """
    def __init__(self, master, run_function, cyclic_run_function):
        """
        Initialize the SingleFileSelector with GUI components.

        """
        self.master = master
        self.run_function = run_function
        self.cyclic_run_function = cyclic_run_function
        self.cyclic_run_var = tk.BooleanVar()
        self.frame = tk.Frame(self.master)
        self.frame.grid(row=0, column=0)

        self.label = tk.Label(self.frame, text="Select file to analyze:")
        self.label.grid(row=0, column=0)

        self.file_entry = tk.Entry(self.frame, width=50)
        self.file_entry.grid(row=0, column=1)

        self.file_button = tk.Button(self.frame, text="Select File", command=self.select_file)
        self.file_button.grid(row=0, column=2)

        self.save_label = tk.Label(self.frame, text="Select folder to save report:")
        self.save_label.grid(row=1, column=0)

        self.save_entry = tk.Entry(self.frame, width=50)
        self.save_entry.grid(row=1, column=1)

        self.save_button = tk.Button(self.frame, text="Select Folder", command=self.select_save_path)
        self.save_button.grid(row=1, column=2)

        self.cyclic_run_checkbox = tk.Checkbutton(
            self.frame,
            text="Enable Cyclic Run Analysis",
            variable=self.cyclic_run_var
        )
        self.cyclic_run_checkbox.grid(row=2, column=0, columnspan=2)

        self.run_button = tk.Button(
            self.frame,
            text="Run",
            command=self.execute_function
        )
        self.run_button.grid(row=3, column=0, columnspan=3)

        self.filepath = ""
        self.savepath = ""

    def select_file(self):
        """
        Open a file dialog to select a single file and update the file entry widget.
        """
        self.filepath = filedialog.askopenfilename()
        if self.filepath:
            self.file_entry.delete(0, tk.END)
            self.file_entry.insert(0, self.filepath)

    def select_save_path(self):
        """
        Open a folder dialog to select a save folder and update the save entry widget.
        """
        self.savepath = filedialog.askdirectory()
        if self.savepath:
            self.save_entry.delete(0, tk.END)
            self.save_entry.insert(0, self.savepath)

    def execute_function(self):
        """
        Execute the appropriate analysis function based on the selected options.

        """
        if self.filepath and self.savepath:
            cyclic_run = self.cyclic_run_var.get()
            self.run_function(self.filepath, self.savepath, cyclic_run=cyclic_run)
        else:
            messagebox.showwarning(
                "Input Missing",
                "Please select a file and a save folder before proceeding."
            )
