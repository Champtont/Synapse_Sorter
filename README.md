
# Synapse Sorter

Synapse Sorter is a Python desktop application that allows you to select questions (with answers and images) from a Microsoft Word (.docx) question bank and copy them into a new, printable exam document, keeping the formatting of the original.

## Features
- Load a .docx file containing a question bank (supports bilingual text and embedded images)
- Select specific questions by number or range (e.g., `20-26, 3, 8`), in the order you want them
- Or load an Excel file with several exam versions (permutations) and generate them all at once
- Keeps the source formatting: fonts, italics, colours, auto-lettered answers and images
- Section titles marked with `*` are never copied
- Optionally removes the red answer marking for a student copy
- Simple, user-friendly graphical interface

## How to Use
1. **Open Synapse Sorter**
2. Enter your desired exam title in the "Exam Title" field
3. Click **LOAD QUESTION BANK** and select your source Word document (.docx)
4. Then either:
   - **Option 1:** enter the question numbers you want (e.g., `20-26, 3, 8`) and click **GENERATE CLEAN EXAM**, or
   - **Option 2:** click **LOAD PERMUTATIONS EXCEL**, select your Excel file, click **GENERATE ALL PERMUTATIONS** and choose a folder. One exam is created per column, named `<Exam Title> - <column name>.docx`
5. Tick **Remove red answer marking** if the exam is for students
6. The app notifies you when the exams are ready, and lists any numbers it could not find

## Permutations Excel Format
- Only the first sheet is read
- Row 1 holds the name of each version, e.g. `Permuta_1`, `Permuta_2`, ...
- Each column below lists the question numbers in the order they should appear

| Permuta_1 | Permuta_2 | Permuta_3 | Permuta_4 |
|-----------|-----------|-----------|-----------|
| 297       | 106       | 108       | 131       |
| 15        | 108       | 63        | 122       |
| ...       | ...       | ...       | ...       |

## Acceptable Formatting for Source Documents
- Each question should start with a number followed by a period or parenthesis, e.g., `1.`, `2)`, etc.
- Answers can be a Word auto-lettered list, or typed as `a.`, `b.`, etc.
- Everything between one question and the next is copied with that question (answers, images, notes)
- Section titles must start with `*` (e.g. `*INNATE IMMUNITY`) so they are left out
- Images should be embedded in the Word document (not linked)

## Notes
- Questions are renumbered from 1 in the order you selected them
- Each question's answer list restarts at `a.`
- The output is a .docx file compatible with Microsoft Word

## Troubleshooting

If some questions are not read or extracted properly, check the following:
- Ensure each question starts with a number and a period or parenthesis (e.g., `1.`, `2)`).
- Make sure every section title starts with `*`, otherwise it is copied with the question before it.
- A number must only be used once in the question bank; the app warns you when it finds duplicates.
- Close the output file in Word before generating it again.

## Development

Set up a virtual environment and install the dependencies:

```bash
python -m venv .venv
.venv\Scripts\python -m pip install -r requirements-dev.txt
```

Run the app from source:

```bash
.venv\Scripts\python main.py
```

Run the tests (no Word needed; test documents are created on the fly):

```bash
.venv\Scripts\python -m pytest
```

To also test against the real question bank and permutations Excel, point the tests at them:

```bash
set SYNAPSE_POOL=C:\path\to\Question_pool_auto.docx && set SYNAPSE_PERMUTAS=C:\path\to\Examen_Midterm_Permutas.xlsx && .venv\Scripts\python -m pytest
```

Build the Windows executable (written to `dist\SynapseSorter.exe`):

```bash
.venv\Scripts\pyinstaller SynapseSorter.spec
```

The code is split so it can be tested without opening the window:
- `synapse_core.py`: reading the question bank and Excel, building exams
- `main.py`: the window
- `tests/`: automated tests

## Tools Used to Build This App

- **Python 3**: Core programming language
- **Tkinter & ttkbootstrap**: For the graphical user interface
- **python-docx**: For reading and writing Word documents
- **openpyxl**: For reading the permutations Excel file
- **pytest**: For automated tests
- **PyInstaller**: For packaging the app into a standalone Windows executable

Developed for Tamara
