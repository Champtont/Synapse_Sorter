import json
import os
import queue
import sys
import threading
import tkinter as tk
from tkinter import PhotoImage, filedialog, messagebox
import ttkbootstrap as tb
from ttkbootstrap.constants import *
try:
    from ttkbootstrap.widgets.scrolled import ScrolledFrame
except ImportError:  # ttkbootstrap < 1.19
    from ttkbootstrap.scrolled import ScrolledFrame

from synapse_core import (ExamHeader, QuestionBank, build_exam, describe_report, parse_number_input,
                          read_permutations, safe_filename)

SETTINGS_PATH = os.path.join(os.environ.get("APPDATA", os.path.expanduser("~")), "SynapseSorter", "settings.json")
FONT = "Segoe UI"
TAB_EXCEL, TAB_NUMBERS = 0, 1

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)

def load_settings(path=None):
    try:
        with open(path or SETTINGS_PATH, encoding="utf-8") as f:
            return json.load(f)
    except (OSError, ValueError):
        return {}

def save_settings(settings, path=None):
    path = path or SETTINGS_PATH
    try:
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=2)
    except OSError:
        pass  # Remembering is a convenience; never block exam generation over it

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
        self.root.geometry("600x860")
        self.root.minsize(520, 500)
        self.bank, self.permutations, self.header = None, None, None
        self.header_entries = {}    # header field index -> Entry
        self.settings = load_settings()
        # The worker thread never touches widgets; it posts messages here instead
        self.events = queue.Queue()

        self.setup_styles()

        # Banner
        banner = tb.Frame(self.root, bootstyle="primary", padding=(28, 18))
        banner.pack(fill="x")
        tb.Label(banner, text="Synapse Sorter", font=(FONT, 20, "bold"), bootstyle="inverse-primary").pack(anchor="w")
        tb.Label(banner, text="Build exams from your question bank", font=(FONT, 10), bootstyle="inverse-primary").pack(anchor="w")

        self.frame = ScrolledFrame(self.root, padding=(20, 16, 20, 8), autohide=True)
        self.frame.configure(style="Page.TFrame")
        self.frame.container.configure(style="Page.TFrame")
        self.frame.container.pack(fill="both", expand=True)

        # 1. Question bank
        card = self.card(1, "Question bank", "The Word document with all the numbered questions.")
        self.load_btn = tb.Button(card, text="Choose question bank…", command=self.load_file, bootstyle="outline-primary")
        self.load_btn.pack(anchor="w", pady=(10, 0))
        self.status = self.status_label(card, "No file selected")

        # 2. Exam header
        card = self.card(2, "Exam header", "Placed at the top of every exam.")
        tip = tb.Frame(card, bootstyle="info", padding=(10, 8))
        tip.pack(fill="x", pady=(10, 0))
        tb.Label(tip, text="In Word, highlight in yellow everything that changes from exam to exam (exam type, date…). "
                           "Only yellow text gets a box here. Yellow “Permuta 1” is filled in for you.",
                 font=(FONT, 9), bootstyle="inverse-info", wraplength=480, justify="left").pack(anchor="w")
        self.header_btn = tb.Button(card, text="Choose exam header…", command=self.choose_header, bootstyle="outline-primary")
        self.header_btn.pack(anchor="w", pady=(10, 0))
        self.header_status = self.status_label(card, "No header selected")
        self.header_fields = tb.Frame(card)
        self.header_fields.pack(fill="x")

        # 3. Questions
        card = self.card(3, "Questions", "Choose which questions go in the exams, and in what order.")
        self.tabs = tb.Notebook(card, bootstyle="primary")
        self.tabs.pack(fill="x", pady=(10, 0))

        excel_tab = tb.Frame(self.tabs, padding=12)
        self.tabs.add(excel_tab, text="  Permutations from Excel  ")
        tb.Label(excel_tab, text="One exam per column. Version name in row 1, question numbers below.",
                 font=(FONT, 9), bootstyle="secondary", wraplength=460).pack(anchor="w")
        self.excel_btn = tb.Button(excel_tab, text="Choose Excel file…", command=self.load_excel, bootstyle="outline-primary")
        self.excel_btn.pack(anchor="w", pady=(10, 0))
        self.excel_status = self.status_label(excel_tab, "No Excel selected")

        numbers_tab = tb.Frame(self.tabs, padding=12)
        self.tabs.add(numbers_tab, text="  Type numbers  ")
        tb.Label(numbers_tab, text="One exam, questions in the order typed. Example: 20-26, 3, 8",
                 font=(FONT, 9), bootstyle="secondary", wraplength=460).pack(anchor="w")
        self.range_ent = tb.Entry(numbers_tab, bootstyle="primary", font=(FONT, 10))
        self.range_ent.pack(fill="x", pady=(10, 0))

        # 4. Create
        card = self.card(4, "Create exams", None)
        tb.Label(card, text="File name", font=(FONT, 10, "bold")).pack(anchor="w", pady=(10, 0))
        self.name_ent = tb.Entry(card, bootstyle="primary", font=(FONT, 10))
        self.name_ent.insert(0, "Immunology Assessment")
        self.name_ent.pack(fill="x", pady=(4, 0))
        tb.Label(card, text="Permutations are saved as “<File name> - Permuta_1.docx”, …",
                 font=(FONT, 9), bootstyle="secondary").pack(anchor="w", pady=(2, 0))

        self.hide_answers = tk.BooleanVar(value=False)
        tb.Checkbutton(card, text="Student copy (remove the red answer marking)", variable=self.hide_answers,
                       bootstyle="primary-round-toggle").pack(anchor="w", pady=(14, 0))

        self.gen_btn = tb.Button(card, text="Create exams", command=self.create, bootstyle="primary", padding=(0, 10))
        self.gen_btn.pack(fill="x", pady=(16, 0))
        self.progress = tb.Progressbar(card, bootstyle="success", value=0)
        self.progress.pack(fill="x", pady=(10, 0))

        # Footer
        footer = tb.Frame(self.frame, style="Page.TFrame")
        footer.pack(fill="x", pady=(0, 8))
        help_link = tb.Label(footer, text="Help & instructions", font=(FONT, 9, "underline"), bootstyle="primary",
                             background=self.page_bg, cursor="hand2")
        help_link.pack(side="left")
        help_link.bind("<Button-1>", lambda _: self.show_help())
        tb.Label(footer, text="Developed for Tamara", font=(FONT, 8, "italic"), bootstyle="secondary",
                 background=self.page_bg).pack(side="right")

        self.buttons = [self.load_btn, self.header_btn, self.excel_btn, self.gen_btn]

        # Reopen the header used last time
        last_header = self.settings.get("header_path")
        if last_header and os.path.exists(last_header):
            self.load_header(last_header, quiet=True)
        elif last_header:
            self.set_status(self.header_status, f"Last header ({os.path.basename(last_header)}) was not found. Choose it again.", "warning")

    # ----------------------------
    # Look
    # ----------------------------

    def setup_styles(self):
        style = self.root.style
        self.page_bg = style.colors.light
        style.configure(".", font=(FONT, 10))
        style.configure("Page.TFrame", background=self.page_bg)
        style.configure("TNotebook.Tab", font=(FONT, 9))

    def card(self, step, title, subtitle):
        """A white panel with a numbered heading, one per step."""
        card = tb.Frame(self.frame, padding=(18, 16))
        card.pack(fill="x", pady=(0, 14))
        head = tb.Frame(card)
        head.pack(fill="x")
        tb.Label(head, text=str(step), font=(FONT, 10, "bold"), bootstyle="inverse-primary",
                 width=2, anchor="center", padding=(0, 2)).pack(side="left")
        tb.Label(head, text=title, font=(FONT, 12, "bold")).pack(side="left", padx=(10, 0))
        if subtitle:
            tb.Label(card, text=subtitle, font=(FONT, 9), bootstyle="secondary", wraplength=500).pack(anchor="w", pady=(6, 0))
        return card

    def status_label(self, parent, text):
        label = tb.Label(parent, font=(FONT, 9), wraplength=480, justify="left")
        label.pack(anchor="w", pady=(6, 0))
        self.set_status(label, text, "idle")
        return label

    def set_status(self, label, text, state):
        icon, style = {"idle": ("○", "secondary"), "success": ("✔", "success"), "warning": ("⚠", "warning")}[state]
        label.config(text=f"{icon}  {text}", bootstyle=style)

    # ----------------------------
    # Actions
    # ----------------------------

    def show_help(self):
        help_msg = (
            "1. Choose your question bank (Word document).\n"
            "2. Choose your exam header (Word document) and fill in the boxes.\n"
            "   In Word, highlight in yellow anything that changes per exam;\n"
            "   only yellow text gets a box. Yellow 'Permuta 1' is filled in\n"
            "   automatically from the Excel.\n"
            "3. Either choose an Excel file with one column per version,\n"
            "   or type the question numbers you want (e.g., 20-26, 3, 8).\n"
            "4. Enter a file name and click Create exams.\n\n"
            "Questions keep the order you give and the formatting of the\n"
            "source document, and are never split across two pages.\n"
            "Titles starting with * are never copied."
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
            self.set_status(self.status, text + f" (numbers used twice: {', '.join(map(str, self.bank.duplicates))})", "warning")
        else:
            self.set_status(self.status, text, "success")

    def choose_header(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            self.load_header(path)

    def load_header(self, path, quiet=False):
        try:
            header = ExamHeader(path)
        except Exception as e:
            if not quiet:
                messagebox.showerror("Could not read header", f"{os.path.basename(path)} could not be read:\n{e}")
            return
        self.header = header
        self.settings["header_path"] = path
        save_settings(self.settings)

        for child in self.header_fields.winfo_children():
            child.destroy()
        self.header_entries = {}
        remembered = self.settings.get("header_values", {})
        for index, original in header.inputs:
            label = original.replace("\t", " ").strip()
            row = tb.Frame(self.header_fields)
            row.pack(fill="x", pady=(8, 0))
            tb.Label(row, text=f"“{label}”  →", width=18, anchor="w", bootstyle="secondary").pack(side="left")
            entry = tb.Entry(row, bootstyle="primary", font=(FONT, 10))
            entry.insert(0, remembered.get(original, original))
            entry.pack(side="left", fill="x", expand=True)
            self.header_entries[index] = entry

        name = os.path.basename(path)
        if not header.fields:
            self.set_status(self.header_status, f"Using {name}, but it has no yellow text, so nothing can be changed.", "warning")
            if not quiet:
                messagebox.showwarning("No yellow text in header",
                                       f"{name} has no text highlighted in yellow, so it will be copied exactly as it is.\n\n"
                                       "To change a part for each exam (e.g. the date), highlight it in yellow in Word, "
                                       "save, and choose the header again.")
        elif not header.has_version:
            self.set_status(self.header_status, f"Using {name}. No yellow “Permuta” text found, so the version will not be printed.", "warning")
        else:
            self.set_status(self.header_status, f"Using {name}. The Permuta number is filled in automatically.", "success")

    def header_values(self):
        values = {i: e.get() for i, e in self.header_entries.items()}
        if self.header:
            self.settings["header_values"] = {self.header.fields[i]["text"]: v for i, v in values.items()}
            save_settings(self.settings)
        return values

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
        self.set_status(self.excel_status, f"{len(self.permutations)} versions: {summary}", "success")

    def _ready(self):
        if not self.bank:
            messagebox.showwarning("No question bank", "Choose a question bank first (step 1).")
            return False
        if not self.header:
            return messagebox.askyesno("No exam header",
                                       "No exam header is loaded (step 2), so the exams will only have the file name as a title.\n\n"
                                       "Create the exams without the header?", icon="warning", default="no")
        return True

    def create(self):
        if self.tabs.index("current") == TAB_EXCEL:
            self.generate_permutations()
        else:
            self.generate()

    def generate(self):
        nums = parse_number_input(self.range_ent.get())
        if not nums:
            messagebox.showwarning("No questions selected", "Type the question numbers you want, e.g. 1-5, 8.")
            return
        if not self._ready(): return
        name = self.name_ent.get()
        save_path = filedialog.asksaveasfilename(defaultextension=".docx", initialfile=safe_filename(name) + ".docx",
                                                 filetypes=[("Word Documents", "*.docx")])
        if save_path:
            self.start([(name, nums, save_path, "")])

    def generate_permutations(self):
        if not self.permutations:
            messagebox.showwarning("No Excel loaded", "Choose a permutations Excel first, or switch to the Type numbers tab.")
            return
        if not self._ready(): return
        folder = filedialog.askdirectory(title="Choose a folder for the exams")
        if not folder: return
        name = self.name_ent.get()
        jobs = [(f"{name} - {version}", nums, os.path.join(folder, safe_filename(f"{name} - {version}") + ".docx"), version)
                for version, nums in self.permutations]
        existing = [os.path.basename(job[2]) for job in jobs if os.path.exists(job[2])]
        if existing and not messagebox.askyesno("Replace files?", "These files already exist and will be replaced:\n" + "\n".join(existing)):
            return
        self.start(jobs)

    def start(self, jobs):
        for b in self.buttons: b.config(state="disabled")
        self.progress["maximum"] = sum(len(job[1]) for job in jobs)
        self.progress["value"] = 0
        threading.Thread(target=self.run_task, args=(jobs, self.hide_answers.get(), self.header, self.header_values()),
                         daemon=True).start()
        self.root.after(100, self.poll_events)

    def run_task(self, jobs, hide_answers, header, header_values):
        done = 0
        results = []
        try:
            for title, nums, path, version in jobs:
                report = build_exam(self.bank, nums, title, path, hide_answers,
                                    on_progress=lambda step, base=done: self.events.put(("progress", base + step)),
                                    header=header, header_values=header_values, version=version)
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
