
# Synapse Sorter

Synapse Sorter is a Python desktop application that allows you to extract, select, and reformat questions (with answers and images) from a Microsoft Word (.docx) question bank into a clean, printable exam document.

## Features
- Load a .docx file containing a question bank (supports bilingual text and embedded images)
- Select specific questions by number or range (e.g., `1-5, 8, 20-26`)
- Automatically restores images and answer formatting
- Generates a new, cleanly formatted .docx exam file with your chosen questions
- Simple, user-friendly graphical interface

## How to Use
1. **Open Synapse Sorter**
2. Enter your desired exam title in the "Exam Title" field
3. Click **LOAD QUESTION BANK** and select your source Word document (.docx)
4. Enter the question numbers you want to include (e.g., `1-10, 15`)
5. Click **GENERATE CLEAN EXAM** and choose where to save your new exam file
6. The app will process your selection and notify you when the new exam is ready

## Acceptable Formatting for Source Documents
- Each question should start with a number followed by a period or parenthesis, e.g., `1.`, `2)`, etc.
- Answers should be labeled with a letter and a period or parenthesis, e.g., `A.`, `B)`, etc.
- Images should be embedded in the Word document (not linked)
- Avoid extra text like "Image alt text" or "AI-generated content" in the question bank
- Example question format:
	```
	1. What is the capital of France?
	A. Paris
	B. London
	C. Berlin
	D. Madrid
	```
- Ranges and individual numbers can be selected in the app (e.g., `1-5, 8, 20-26`)

## Notes
- The app works best with question banks that follow the above formatting
- Both text and images are preserved in the output
- The output is a .docx file compatible with Microsoft Word

## Troubleshooting

If some questions are not read or extracted properly, check the following:
- Ensure each question starts with a number and a period or parenthesis (e.g., `1.`, `2)`).
- Answers must be labeled with a letter and a period or parenthesis (e.g., `A.`, `B)`).
- Avoid extra or unexpected text (like "Image alt text" or "AI-generated content") between questions and answers.
- Images should be embedded, not linked.
- The app may not recognize questions with unusual formatting, missing numbers, or inconsistent answer labels.
- For best results, use clear and consistent formatting as described above.

If you encounter issues, try editing your source document to match the recommended format and reload it in the app.

## Tools Used to Build This App

- **Python 3**: Core programming language
- **Tkinter & ttkbootstrap**: For the graphical user interface
- **python-docx**: For creating and editing Word documents
- **docx2python**: For extracting text and images from Word documents
- **PyInstaller**: For packaging the app into a standalone Windows executable

Developed for Tamara
