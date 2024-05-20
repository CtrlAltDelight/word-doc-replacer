import docx
import re
import sys
import tkinter as tk
from tkinter import ttk
from tkinter import filedialog

class App(tk.Frame):
    def __init__(self, master=None):
        super().__init__(master)
        self.pack()

        self.entrythingy = tk.Entry()
        self.entrythingy.pack()
        self.contents = tk.StringVar()
        self.contents.set("this is a variable")
        self.entrythingy["textvariable"] = self.contents
        self.entrythingy.bind('<Key-Return>', self.print_contents)

        self.create_widgets()


    def create_widgets(self):
        self.selectButton = tk.Button(text="Select File", command=self.select_file)
        self.selectButton.pack()

    def print_contents(self, event):
        print("Hi. The current entry content is:", self.contents.get())

    def select_file(self):
        file = filedialog.askopenfile(mode='r')
        if file:
            print("File selected: ", file.name)

if __name__ == "__main__":
    root = tk.Tk()
    myapp = App(master=root)
    myapp.master.title("My App")
    myapp.master.maxsize(1000, 400)
    myapp.mainloop()


"""
def prompt_user(question):
    print(question)
    answer = input()
    if answer.lower() == 'quit':
        print("Goodbye!")
        sys.exit()
    return answer

# Create a Tkinter root widget
root = tk.Tk(screenName="Format Code Replacer")
frame = ttk.Frame(root, padding=10)
frame.grid()
ttk.Label(frame, text="Hello world").grid(row=0, column=0)
ttk.Button(frame, text="Quit", command=root.destroy).grid(row=0, column=1)
root.mainloop()
#root.withdraw()  # Hide the main window

# Ask the user to select a file
file_path = filedialog.askopenfilename(filetypes=[("Word Documents", "*.docx")])
if not file_path:
    print("No file selected.")
    wait = input("Press enter to exit program.")
    sys.exit()

# Open the document
try:
    doc = docx.Document(file_path)
except Exception as e:
    print(f"Error opening document: {e}")
    wait = input("Press enter to exit program.")
    sys.exit()

# Replace codes in the document
answered_codes = dict() # keeps track of answered codes so repeated codes are not asked again
for paragraph in doc.paragraphs:
"""
    #codes = re.findall(r"\{[^\}]+\}", paragraph.text)
"""
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
        
        original_text = paragraph.text
        paragraph.clear()
        new_text = original_text.replace("{" + code + "}", answer)
        paragraph.add_run(new_text)
        print(answered_codes)

# Prompt user to save the modified document
try:
    output_file_path = filedialog.asksaveasfilename(defaultextension=".docx", filetypes=[("Word Documents", "*.docx")])
    if output_file_path:
        doc.save(output_file_path)
        print("The program has finished. Please check the modified document.")
    else:
        print("No save location selected. Exiting without saving.")
except Exception as e:
    print(f"Error saving document: {e}")

"""
