from pbix.PBIFile import PBIFile
import tkinter as tk
import sys


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

        label_path = tk.Label(frame_inputs, text="Filepath", anchor='w').grid(row=0, padx=padx, pady=pady)
        label_old = tk.Label(frame_inputs, text="Old Measure", anchor='w').grid(row=1, padx=padx, pady=5)
        label_new = tk.Label(frame_inputs, text="New Measure", anchor='w').grid(row=2, padx=padx, pady=5)

        self.entry_path = tk.Entry(frame_inputs)
        self.entry_old = tk.Entry(frame_inputs)
        self.entry_new = tk.Entry(frame_inputs)

        self.entry_path.grid(row=0, column=1, sticky='ew', padx=padx, pady=5)
        self.entry_old.grid(row=1, column=1, sticky='ew', padx=padx, pady=5)
        self.entry_new.grid(row=2, column=1, sticky='ew', padx=padx, pady=5)

        tk.Button(frame_buttons, text='Quit', command=self.quit).grid(row=3, column=0, sticky="w", pady=10, padx=10)
        tk.Button(frame_buttons, text='Update Measures', command=self.update_measures).grid(row=3, column=1, sticky="w", pady=10, padx=10)

        self.text = tk.Text(frame_output, wrap="word")
        self.text.pack(side='left', fill='both', padx=padx, pady=pady)
        self.text.tag_configure("stderr", foreground="#b22222")

        sys.stdout = TextRedirector(self.text, "stdout")
        sys.stderr = TextRedirector(self.text, "stderr")

    def _get_inputs(self):
        path = self.entry_path.get()
        old = self.entry_old.get()
        new = self.entry_new.get()
        return path, old, new


    def update_measures(self):
        path, old, new = self._get_inputs()
        report = PBIFile(path)
        report.update_measures(old, new)
        if report.updated > 0:
            report.write_file()


class TextRedirector(object):
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