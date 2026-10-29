import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import PhotoImage, filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *

from synapse_core import (QuestionBank, build_exam, describe_report, parse_number_input,
                          read_permutations, safe_filename)

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

# ----------------------------
# UI Class
# ----------------------------

class SynapseSorterApp:
    def __init__(self, root):
        self.root = root
        icon_file = resource_path("icon.ico")  # Make sure icon.ico exists
        if os.path.exists(icon_file):
            try:
                self.root.iconbitmap(icon_file)  # Windows
            except tk.TclError:
                pass
        else:
            # Cross-platform fallback using .png
            try:
                photo = PhotoImage(file=resource_path("icon.png"))
                self.root.tk.call('wm', 'iconphoto', self.root._w, photo)
            except Exception as e:
                print("Failed to load icon:", e)
            print("Icon file not found at", icon_file)

        self.root.title("Synapse Sorter")
        self.root.geometry("550x800")
        self.bank, self.permutations = None, None
        # The worker thread never touches widgets; it posts messages here instead
        self.events = queue.Queue()

        self.frame = tb.Frame(self.root, padding=30)
        self.frame.pack(fill="both", expand=True)

        tb.Label(self.frame, text="SYNAPSE SORTER", font=("Helvetica", 20, "bold"), bootstyle="primary").pack(pady=(0, 25))

        tb.Label(self.frame, text="Exam Title:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        self.title_ent = tb.Entry(self.frame, bootstyle="primary")
        self.title_ent.insert(0, "Immunology Assessment")
        self.title_ent.pack(fill="x", pady=(5, 20))

        self.load_btn = tb.Button(self.frame, text="LOAD QUESTION BANK", command=self.load_file, bootstyle="outline-primary")
        self.load_btn.pack(fill="x", pady=5)
        self.status = tb.Label(self.frame, text="No file selected", font=("Helvetica", 9), bootstyle="secondary")
        self.status.pack(pady=(0, 20))

        tb.Label(self.frame, text="Option 1 - Type Question Numbers:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        tb.Label(self.frame, text="Questions appear in the order typed. Example: 20-26, 3, 8", font=("Helvetica", 8), bootstyle="secondary").pack(anchor="w")
        self.range_ent = tb.Entry(self.frame, bootstyle="primary")
        self.range_ent.pack(fill="x", pady=5)
        self.gen_btn = tb.Button(self.frame, text="GENERATE CLEAN EXAM", command=self.generate, bootstyle="primary")
        self.gen_btn.pack(fill="x", pady=(5, 20))

        tb.Label(self.frame, text="Option 2 - Permutations From Excel:", font=("Helvetica", 10, "bold")).pack(anchor="w")
        tb.Label(self.frame, text="One column per version, version name in row 1", font=("Helvetica", 8), bootstyle="secondary").pack(anchor="w")
        self.excel_btn = tb.Button(self.frame, text="LOAD PERMUTATIONS EXCEL", command=self.load_excel, bootstyle="outline-primary")
        self.excel_btn.pack(fill="x", pady=5)
        self.excel_status = tb.Label(self.frame, text="No Excel selected", font=("Helvetica", 9), bootstyle="secondary")
        self.excel_status.pack()
        self.perm_btn = tb.Button(self.frame, text="GENERATE ALL PERMUTATIONS", command=self.generate_permutations, bootstyle="primary")
        self.perm_btn.pack(fill="x", pady=(5, 20))

        self.hide_answers = tk.BooleanVar(value=False)
        tb.Checkbutton(self.frame, text="Remove red answer marking (student copy)", variable=self.hide_answers, bootstyle="primary").pack(anchor="w")

        self.progress = tb.Progressbar(self.frame, bootstyle="primary-striped", value=0)
        self.progress.pack(fill="x", pady=20)

        tb.Button(self.frame, text="HELP / INSTRUCTIONS", command=self.show_help, bootstyle="link").pack(pady=5)
        tb.Label(self.frame, text="Developed for Tamara", font=("Helvetica", 8, "italic"), bootstyle="secondary").pack(side="bottom", pady=10)

        self.buttons = [self.load_btn, self.gen_btn, self.excel_btn, self.perm_btn]

    def show_help(self):
        help_msg = (
            "1. Enter an Exam Title.\n"
            "2. Click LOAD to select your source Word document.\n"
            "3. Either type the question numbers you want (e.g., 20-26, 3, 8)\n"
            "   and click GENERATE CLEAN EXAM,\n"
            "   or load an Excel file with one column per version and\n"
            "   click GENERATE ALL PERMUTATIONS.\n\n"
            "Questions keep the order you give and the formatting of the\n"
            "source document. Titles starting with * are never copied."
        )
        messagebox.showinfo("Synapse Sorter Help", help_msg)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if not path: return
        try:
            self.bank = QuestionBank(path)
        except Exception as e:
            messagebox.showerror("Could not read document", f"{os.path.basename(path)} could not be read:\n{e}")
            return
        text = f"Loaded {len(self.bank)} questions from {os.path.basename(path)}"
        if self.bank.duplicates:
            text += f" (numbers used twice: {', '.join(map(str, self.bank.duplicates))})"
        self.status.config(text=text, bootstyle="warning" if self.bank.duplicates else "success")

    def load_excel(self):
        path = filedialog.askopenfilename(filetypes=[("Excel Workbooks", "*.xlsx")])
        if not path: return
        try:
            self.permutations = read_permutations(path)
        except Exception as e:
            messagebox.showerror("Could not read Excel", f"{os.path.basename(path)} could not be read:\n{e}")
            return
        if not self.permutations:
            messagebox.showwarning("Empty Excel", "No question numbers were found in the first sheet.")
            return
        summary = ", ".join(f"{name} ({len(nums)})" for name, nums in self.permutations)
        self.excel_status.config(text=f"{len(self.permutations)} versions: {summary}", bootstyle="success", wraplength=480)

    def _ready(self):
        if not self.bank:
            messagebox.showwarning("No question bank", "Load a question bank first.")
            return False
        return True

    def generate(self):
        if not self._ready(): return
        nums = parse_number_input(self.range_ent.get())
        if not nums:
            messagebox.showwarning("No questions selected", "Enter question numbers, e.g. 1-5, 8.")
            return
        title = self.title_ent.get()
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=safe_filename(title) + ".docx",
                                                 filetypes=[("Word Documents", "*.docx")])
        if save_path:
            self.start([(title, nums, save_path)])

    def generate_permutations(self):
        if not self._ready(): return
        if not self.permutations:
            messagebox.showwarning("No Excel loaded", "Load a permutations Excel first.")
            return
        folder = filedialog.askdirectory(title="Choose a folder for the exams")
        if not folder: return
        title = self.title_ent.get()
        jobs = [(f"{title} - {name}", nums, os.path.join(folder, safe_filename(f"{title} - {name}") + ".docx"))
                for name, nums in self.permutations]
        existing = [os.path.basename(p) for _, _, p in jobs if os.path.exists(p)]
        if existing and not messagebox.askyesno("Replace files?", "These files already exist and will be replaced:\n" + "\n".join(existing)):
            return
        self.start(jobs)

    def start(self, jobs):
        for b in self.buttons: b.config(state="disabled")
        self.progress["maximum"] = sum(len(nums) for _, nums, _ in jobs)
        self.progress["value"] = 0
        threading.Thread(target=self.run_task, args=(jobs, self.hide_answers.get()), daemon=True).start()
        self.root.after(100, self.poll_events)

    def run_task(self, jobs, hide_answers):
        done = 0
        results = []
        try:
            for title, nums, path in jobs:
                report = build_exam(self.bank, nums, title, path, hide_answers,
                                    on_progress=lambda step, base=done: self.events.put(("progress", base + step)))
                done += len(nums)
                results.append((path, report))
            self.events.put(("done", results))
        except PermissionError as e:
            self.events.put(("error", f"Could not save {e.filename}.\nIf it is open in Word, close it and try again."))
        except Exception as e:
            self.events.put(("error", str(e)))

    def poll_events(self):
        try:
            while True:
                kind, payload = self.events.get_nowait()
                if kind == "progress":
                    self.progress["value"] = payload
                    continue
                for b in self.buttons: b.config(state="normal")
                if kind == "error":
                    messagebox.showerror("Generation failed", payload)
                else:
                    self.show_results(payload)
                return
        except queue.Empty:
            self.root.after(100, self.poll_events)

    def show_results(self, results):
        lines = []
        for path, report in results:
            lines.append(os.path.basename(path))
            details = describe_report(report)
            if details: lines.append("   " + details.replace("\n", "\n   "))
        if any(describe_report(r) for _, r in results):
            messagebox.showwarning("Exam Generated (check notes)", "\n".join(lines))
        else:
            messagebox.showinfo("Success", "Exam Generated!\n\n" + "\n".join(lines))

if __name__ == "__main__":
    root_win = tb.Window(themename="flatly")
    app = SynapseSorterApp(root_win)
    root_win.mainloop()
