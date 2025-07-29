import tkinter as tk
from tkinter import filedialog, messagebox
from pathlib import Path
import threading
import itertools
import time

from processors.federal_processor import process_federal_files
from processors.state_processor import attach_w2_to_stfcs
from processors.combiner import combine_state_files
from processors.envelope_creator import create_envelope_docs

class SortFilesApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Sort & Edit Files Software")
        self.root.geometry("450x300")
        self.work_dir = None
        self.spinner_running = False

        self.label = tk.Label(root, text="Please select your work directory", wraplength=400)
        self.label.pack(pady=10)

        self.select_btn = tk.Button(root, text="Select Work Directory", command=self.select_directory)
        self.select_btn.pack(pady=5)

        self.process_btn = tk.Button(root, text="Start Processing", command=self.start_workflow_thread, state=tk.DISABLED)
        self.process_btn.pack(pady=15)

        self.spinner_label = tk.Label(root, text="", font=("Courier", 16))
        self.spinner_label.pack(pady=10)

    def select_directory(self):
        folder_selected = filedialog.askdirectory()
        if folder_selected:
            self.work_dir = Path(folder_selected)
            self.label.config(text=f"Selected: {self.work_dir}")
            self.process_btn.config(state=tk.NORMAL)

    def start_workflow_thread(self):
        self.process_btn.config(state=tk.DISABLED)
        self.select_btn.config(state=tk.DISABLED)
        self.spinner_running = True
        threading.Thread(target=self.animate_spinner, daemon=True).start()
        threading.Thread(target=self.run_workflow, daemon=True).start()

    def animate_spinner(self):
        spinner_cycle = itertools.cycle(['|', '/', '-', '\\'])
        while self.spinner_running:
            self.spinner_label.config(text=next(spinner_cycle) + " Processing...")
            time.sleep(0.1)

    def run_workflow(self):
        try:
            print(f"📁 Starting workflow with work directory: {self.work_dir}")
            company = self.work_dir / 'company'
            w2 = self.work_dir / 'W2'
            output_data = self.work_dir / 'output_data'
            federal = self.work_dir / 'federal'
            state = self.work_dir / 'state'
            combined = self.work_dir / 'combined'
            envelopes = self.work_dir / 'envelopes'

            people_dirs = list(output_data.glob('peopleinput*/docs'))
            print(f"📁 Found {len(people_dirs)} peopleinput*/docs directories")

            print("📄 Running process_federal_files")
            process_federal_files(company, federal)

            print("📄 Running attach_w2_to_stfcs")
            attach_w2_to_stfcs(company, w2, state)

            print("📄 Running combine_state_files")
            combined_info = combine_state_files(state, combined)
            print(f"✅ Created {len(combined_info)} combined PDFs")

            print("📄 Running create_envelope_docs")
            create_envelope_docs(combined_info, people_dirs, envelopes)

            messagebox.showinfo("Done", "All processing complete!")
        except Exception as e:
            print(f"❌ Workflow error: {e}")
            messagebox.showerror("Error", f"An error occurred:\n{e}")
        finally:
            self.spinner_running = False
            self.spinner_label.config(text="")
            self.select_btn.config(state=tk.NORMAL)
            self.process_btn.config(state=tk.NORMAL)

if __name__ == '__main__':
    root = tk.Tk()
    app = SortFilesApp(root)
    root.mainloop()
