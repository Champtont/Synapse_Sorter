import os
import re
import threading
import tempfile
import tkinter as tk
from tkinter import ttk, filedialog, messagebox

from docx import Document
from docx.shared import Inches, Pt
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx2python import docx2python

# Attempt to load Drag and Drop support if available
try:
    from tkinterdnd2 import TkinterDnD, DND_FILES
    DND_AVAILABLE = True
except ImportError:
    DND_AVAILABLE = False

# ----------------------------
# Utilities
# ----------------------------

def parse_number_input(text):
    """Parses input like '1, 5-10, 20' into a sorted list of integers."""
    numbers = set()
    for part in text.split(","):
        part = part.strip()
        if not part: continue
        if "-" in part:
            try:
                start, end = part.split("-")
                numbers.update(range(int(start), int(end) + 1))
            except ValueError: continue
        else:
            try: numbers.add(int(part))
            except ValueError: continue
    return sorted(numbers)

def normalize_answer_text(text):
    """
    Standardizes answer letters to 'A. Text' format.
    Fixes 'A)', 'a.', and inconsistent spacing [cite: 59-63, 165-169].
    """
    match = re.match(r'^([A-Ea-e])[\.\)]\s*(.*)', text)
    if match:
        letter = match.group(1).upper()
        content = match.group(2).strip()
        return f"{letter}. {content}"
    return text.strip()

# ----------------------------
# Document Processing
# ----------------------------

def process_document(file_path):
    """Extracts questions, bilingual text, and images from the source doc, preserving images."""
    with docx2python(file_path) as doc:
        text = doc.text
        images = doc.images 

    # Map images to their filenames for consistent lookup
    image_lookup = {os.path.basename(k): v for k, v in images.items()}

    questions = {}
    current_q = None
    
    # Matches '1. (optional metadata) text' and removes metadata like (mid25) or (Final25)
    q_pattern = re.compile(r'^(\d+)\.?\s*(?:\([^\)]*\))?\s*(.*)')
    ans_pattern = re.compile(r'^([A-Ea-e])[\.\)]')

    for line in text.splitlines():
        line = line.strip()
        if not line: 
            continue

        # Detect new question number
        q_match = q_pattern.match(line)
        if q_match:
            current_q = int(q_match.group(1))
            questions[current_q] = {
                "header_parts": [q_match.group(2).strip()],
                "elements": [],
                "found_first_answer": False
            }
            continue

        if current_q is not None:
            # Step 1: Check for image placeholders (unchanged)
            img_match = re.search(r'----(.*?image\d+\.\w+)----', line)
            if img_match:
                img_filename = os.path.basename(img_match.group(1))
                questions[current_q]["elements"].append({
                    "type": "image",
                    "name": img_filename
                })
                continue

            # Step 2: Skip trash/placeholder text
            trash_patterns = ["Image alt text", "AI-generated content", "----image"]
            if any(p in line for p in trash_patterns):
                continue

            # Step 3: Decide if line is part of Header or Answer
            if ans_pattern.match(line):
                questions[current_q]["found_first_answer"] = True
                questions[current_q]["elements"].append({
                    "type": "text",
                    "content": line
                })
            elif not questions[current_q]["found_first_answer"]:
                # Lines before first answer are part of the header
                questions[current_q]["header_parts"].append(line)
            else:
                # Lines after first answer
                questions[current_q]["elements"].append({
                    "type": "text",
                    "content": line
                })

    # Combine multi-line headers and remove empty/placeholder strings
    for q_id in questions:
        clean_parts = [p.strip() for p in questions[q_id]["header_parts"] if p.strip() and p != "?"]
        questions[q_id]["header"] = " ".join(clean_parts) if clean_parts else "Question Text Missing"

    return questions, image_lookup

# ----------------------------
# Document Creation
# ----------------------------

def create_new_test(output_path, title, selected_nums, questions, image_dict, progress_cb):
    """Builds the final Word document with the requested formatting."""
    doc = Document()

    # Apply Global Font Settings
    style = doc.styles['Normal']
    style.font.name = 'Times New Roman'
    style.font.size = Pt(11)

    # Center-aligned Title
    title_p = doc.add_heading(title, level=0)
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER

    with tempfile.TemporaryDirectory() as tmp:
        new_num = 1
        for old_num in selected_nums:
            if old_num not in questions:
                progress_cb()
                continue

            q = questions[old_num]
            
            # Question Header (Bold)
            header_p = doc.add_paragraph()
            header_p.paragraph_format.space_before = Pt(12)
            header_p.paragraph_format.keep_with_next = True
            run = header_p.add_run(f"{new_num}. {q['header']}")
            run.bold = True

            for elem in q["elements"]:
                if elem["type"] == "text":
                    # Normalize answer choices (A. Text) [cite: 59-63]
                    formatted = normalize_answer_text(elem["content"])
                    p = doc.add_paragraph(formatted)
                    p.paragraph_format.left_indent = Inches(0.3)
                    p.paragraph_format.space_after = Pt(2)
                    p.paragraph_format.keep_with_next = True 
                
                elif elem["type"] == "image":
                    img_name = elem["name"]
                    if img_name in image_dict:
                        img_path = os.path.join(tmp, img_name)
                        with open(img_path, "wb") as f:
                            f.write(image_dict[img_name])
                        
                        img_p = doc.add_paragraph()
                        img_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
                        run = img_p.add_run()
                        run.add_picture(img_path, width=Inches(2.5))
                        img_p.paragraph_format.space_before = Pt(6)
                        img_p.paragraph_format.space_after = Pt(6)

            # Add spacing between questions
            doc.add_paragraph().paragraph_format.space_after = Pt(6)
            new_num += 1
            progress_cb()

    doc.save(output_path)

# ----------------------------
# UI Logic (Synapse Sorter App)
# ----------------------------

class SynapseSorterApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Synapse Sorter")
        self.root.geometry("550x500")
        self.questions = None
        self.image_dict = None
        self.build_ui()

    def build_ui(self):
        main_frame = ttk.Frame(self.root, padding=25)
        main_frame.pack(fill="both", expand=True)

        ttk.Label(main_frame, text="Synapse Sorter", font=("Segoe UI", 18, "bold")).pack(pady=(0, 20))

        ttk.Label(main_frame, text="1. Enter Exam Title:").pack(anchor="w")
        self.title_entry = ttk.Entry(main_frame)
        self.title_entry.insert(0, "Immunology Assessment")
        self.title_entry.pack(fill="x", pady=(0, 15))

        self.load_btn = ttk.Button(main_frame, text="2. Load Question Bank (.docx)", command=self.load_file)
        self.load_btn.pack(fill="x", pady=5)

        self.status = ttk.Label(main_frame, text="Status: Waiting for file...", foreground="gray")
        self.status.pack(pady=(0, 15))

        ttk.Label(main_frame, text="3. Question Numbers (e.g. 1-5, 8, 20-26):").pack(anchor="w")
        self.range_entry = ttk.Entry(main_frame)
        self.range_entry.pack(fill="x", pady=(0, 15))

        self.progress = ttk.Progressbar(main_frame, mode="determinate")
        self.progress.pack(fill="x", pady=10)

        self.gen_btn = ttk.Button(main_frame, text="4. Generate Clean Exam", command=self.generate)
        self.gen_btn.pack(fill="x", pady=10)

    def load_file(self):
        path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
        if path:
            try:
                self.questions, self.image_dict = process_document(path)
                self.status.config(text=f"Status: {len(self.questions)} questions loaded.", foreground="green")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to load document: {e}")

    def generate(self):
        if not self.questions:
            messagebox.showwarning("Synapse Sorter", "Please load a source file first.")
            return
        
        try:
            nums = parse_number_input(self.range_entry.get())
            if not nums: raise ValueError
        except:
            messagebox.showerror("Error", "Invalid range format (Example: 1-5, 8).")
            return

        save_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word", "*.docx")])
        if save_path:
            self.progress["maximum"] = len(nums)
            self.progress["value"] = 0
            threading.Thread(target=self.run_task, args=(save_path, self.title_entry.get(), nums), daemon=True).start()

    def run_task(self, path, title, nums):
        def update_progress():
            self.progress["value"] += 1
            self.root.update_idletasks()
        
        try:
            create_new_test(path, title, nums, self.questions, self.image_dict, update_progress)
            messagebox.showinfo("Success", "Exam generated with consistent formatting and images!")
        except Exception as e:
            messagebox.showerror("Error", f"An error occurred during generation: {e}")

if __name__ == "__main__":
    # Launch with Drag and Drop support if library is present
    root = TkinterDnD.Tk() if DND_AVAILABLE else tk.Tk()
    app = SynapseSorterApp(root)
    root.mainloop()