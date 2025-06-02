# main.py

import os
import threading
from pathlib import Path
import tkinter as tk
from tkinter import filedialog, messagebox

from modernizer.parser import parse_legacy_docx_by_sequence
from modernizer.populate_template import populate_template

# Adjust this if your template lives somewhere else:
TEMPLATE_PATH = Path(__file__).parent / "templates" / "new_format_template.docx"

def process_folder(input_folder: Path, status_var: tk.StringVar):
    """
    For every .docx in input_folder, parse + populate into Downloads.
    Updates status_var as it goes.
    """
    downloads_dir = Path.home() / "Downloads"
    downloads_dir.mkdir(exist_ok=True)  # just in case

    docx_files = list(input_folder.glob("*.docx"))
    if not docx_files:
        messagebox.showwarning("No .docx Found", f"No .docx files in {input_folder}")
        return

    for idx, legacy_file in enumerate(docx_files, 1):
        status_var.set(f"({idx}/{len(docx_files)}) Processing {legacy_file.name} ...")
        try:
            parsed = parse_legacy_docx_by_sequence(legacy_file)
            output_path = downloads_dir / legacy_file.name
            populate_template(parsed, TEMPLATE_PATH, output_path)
        except Exception as e:
            messagebox.showerror(
                "Error",
                f"Failed on {legacy_file.name}:\n{e}"
            )
            return

    status_var.set("All done!")
    messagebox.showinfo("Finished", f"All {len(docx_files)} files processed into Downloads.")


class App(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title("Doc Modernizer")
        self.geometry("500x200")
        self.resizable(False, False)

        self.selected_dir: Path = None

        # Widgets:
        self.label = tk.Label(self, text="Select a folder of legacy .docx files:")
        self.label.pack(pady=(20, 5))

        self.dir_var = tk.StringVar(value="(no folder selected)")
        self.dir_label = tk.Label(self, textvariable=self.dir_var, fg="blue")
        self.dir_label.pack()

        self.select_btn = tk.Button(self, text="Browse…", width=12, command=self.browse_folder)
        self.select_btn.pack(pady=10)

        self.process_btn = tk.Button(self, text="Process All → Downloads", width=20, command=self.on_process)
        self.process_btn.pack(pady=(5, 10))

        self.status_var = tk.StringVar(value="Idle")
        self.status_label = tk.Label(self, textvariable=self.status_var, fg="green")
        self.status_label.pack(pady=(5, 0))

    def browse_folder(self):
        folder = filedialog.askdirectory(title="Select Folder with Legacy .docx")
        if folder:
            self.selected_dir = Path(folder)
            self.dir_var.set(str(self.selected_dir))
            self.status_var.set("Ready to process.")

    def on_process(self):
        if not self.selected_dir or not self.selected_dir.exists():
            messagebox.showwarning("No Folder", "Please select a valid folder first.")
            return

        # Disable buttons while processing, to prevent double‐clicks
        self.select_btn.config(state=tk.DISABLED)
        self.process_btn.config(state=tk.DISABLED)

        # Run processing in a background thread so the UI doesn't freeze
        threading.Thread(
            target=self._run_processing,
            daemon=True
        ).start()

    def _run_processing(self):
        try:
            process_folder(self.selected_dir, self.status_var)
        finally:
            # Re‐enable buttons when done (or on error)
            self.select_btn.config(state=tk.NORMAL)
            self.process_btn.config(state=tk.NORMAL)


if __name__ == "__main__":
    # Ensure tkinter is using a theme on Windows
    try:
        import ctypes
        ctypes.windll.shcore.SetProcessDpiAwareness(1)
    except Exception:
        pass

    app = App()
    app.mainloop()
