import multiprocessing
import os
import sys
import tkinter as tk
from pathlib import Path

from pbix.report import Report


class App(tk.Tk):
    def __init__(self):
        tk.Tk.__init__(self)
        self.geometry("500x300")
        padx = 10
        pady = 10

        frame_inputs = tk.Frame(self)
        frame_buttons = tk.Frame(self)
        frame_output = tk.Frame(self)

        frame_inputs.pack(side="top", fill="x")
        frame_inputs.grid_columnconfigure(1, weight=1)
        frame_buttons.pack(side="top", fill="x")
        frame_output.pack(side="top", fill="x")

        tk.Label(frame_inputs, text="Filepath", anchor="w").grid(
            row=0, padx=padx, pady=pady
        )
        tk.Label(frame_inputs, text="Old Measure", anchor="w").grid(
            row=1, padx=padx, pady=5
        )
        tk.Label(frame_inputs, text="New Measure", anchor="w").grid(
            row=2, padx=padx, pady=5
        )
        tk.Label(frame_inputs, text="Model", anchor="w").grid(row=3, padx=padx, pady=5)

        self.entry_path = tk.Entry(frame_inputs)
        self.entry_old = tk.Entry(frame_inputs)
        self.entry_new = tk.Entry(frame_inputs)
        self.entry_model = tk.Entry(frame_inputs)
        self.entry_model.insert(0, "Model.pbix")

        self.entry_path.grid(row=0, column=1, sticky="ew", padx=padx, pady=5)
        self.entry_old.grid(row=1, column=1, sticky="ew", padx=padx, pady=5)
        self.entry_new.grid(row=2, column=1, sticky="ew", padx=padx, pady=5)
        self.entry_model.grid(row=3, column=1, sticky="ew", padx=padx, pady=5)

        tk.Button(frame_buttons, text="Quit", command=self.quit).grid(
            row=4, column=0, sticky="w", pady=10, padx=10
        )
        tk.Button(
            frame_buttons, text="Update Measures", command=self.replace_fields
        ).grid(row=4, column=1, sticky="w", pady=10, padx=10)

        self.text = tk.Text(frame_output, wrap="word")
        self.text.pack(side="left", fill="both", padx=padx, pady=pady)
        self.text.tag_configure("stderr", foreground="#b22222")

        sys.stdout = TextRedirector(self.text, "stdout")
        sys.stderr = TextRedirector(self.text, "stderr")

    def _get_inputs(self):
        path = self.entry_path.get()
        old = self.entry_old.get()
        new = self.entry_new.get()
        model = self.entry_model.get()
        return path, old, new, model

    @staticmethod
    def replace_field_in_report(path, old, new):
        """Replace field in a single pbix file."""
        report = Report(path)
        report.update_fields(old, new)
        if report.updated > 0:
            report.write_file()

    def replace_field_in_all_reports(self, path, old, new, model):
        """Replace field in all pbix files in specified directory and all child directories."""
        pool = multiprocessing.Pool()
        for subdir, _, files in os.walk(path):
            pool.starmap(
                self.replace_field_in_report,
                [
                    (os.path.join(subdir, file), old, new)
                    for file in files
                    if file.split(".")[-1] == "pbix" and file != model
                ],
            )

    def replace_fields(self):
        """Replace field fuction."""
        path, old, new, model = self._get_inputs()
        path = Path(path)
        if path.is_file():
            self.replace_field_in_report(path, old, new)
        else:
            self.replace_field_in_all_reports(path, old, new, model)


class TextRedirector:
    def __init__(self, widget, tag="stdout"):
        self.widget = widget
        self.tag = tag

    def write(self, str):
        self.widget.configure(state="normal")
        self.widget.insert("end", str, (self.tag,))
        self.widget.configure(state="disabled")


if __name__ == "__main__":
    app = App()
    app.mainloop()
