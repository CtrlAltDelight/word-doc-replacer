import docx
import re
import sys
import tkinter as tk
from tkinter import filedialog

def prompt_user(question):
    print(question)
    answer = input()
    if answer.lower() == 'quit':
        print("Goodbye!")
        sys.exit()
    return answer

# Create a Tkinter root widget
root = tk.Tk()
root.withdraw()  # Hide the main window

# Ask the user to select a file
file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
if not file_path:
    print("No file selected.")
    wait = input("Press enter to exit program.")
    sys.exit()

try:
    doc = docx.Document(file_path)
except Exception as e:
    print(f"Error opening document: {e}")
    wait = input("Press enter to exit program.")
    sys.exit()

answered_codes = dict() # keeps track of answered codes so repeated codes are not asked again
for para in doc.paragraphs:
    codes = re.findall(r"\{[^\}]+\}", para.text)
    for code in codes:
        code = code[1:-1]
        parts = code.split("|")
        question = parts[0]
        default = parts[1] if len(parts) > 1 else ""
        
        answer = ""
        if question in answered_codes:
            answer = answered_codes[question]
        else:
            answer = prompt_user(question)
            if answer == "":
                answer = default
            answered_codes[question] = answer
        
        original_text = para.text
        para.clear()
        new_text = original_text.replace("{" + code + "}", answer)
        para.add_run(new_text)
        print(answered_codes)

try:
    output_file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if output_file_path:
        doc.save(output_file_path)
        print("The program has finished. Please check the modified document.")
    else:
        print("No save location selected. Exiting without saving.")
except Exception as e:
    print(f"Error saving document: {e}")

