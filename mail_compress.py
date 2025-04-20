import tkinter as tk
from tkinter import filedialog
import re

def process_file(name, file_path):
    # Read the file content and split it into lines
    with open(file_path, 'r', encoding='utf-8-sig') as file:
        content = file.read()

    input_lines = content.splitlines()

    # Constructing the subject
    subject = input_lines[9].replace("Tên file đường dẫn :", '') + input_lines[10] + \
             f'/{input_lines[11].replace("Địa điểm phân công :", "")}' + \
             f'/{input_lines[12].replace("Ngày phân công : ", "")}' + \
             f'/{name} ' + "HOÀN THÀNH"

    # Constructing the body
    body = input_lines[13] + '\n' + input_lines[14] + '\n' + input_lines[15]

    # Clean the subject and body
    subject = re.sub(r'\s+', ' ', subject).strip()
    body = re.sub(r'\s+', ' ', body).strip()

    return subject, body

def browse_file():
    # Open a file dialog to select a file
    file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
    if file_path:
        # Get the name from the input field and process the file
        name = name_entry.get()
        subject, body = process_file(name, file_path)
        
        # Display the subject and body in the result text boxes
        subject_text.delete(1.0, tk.END)
        subject_text.insert(tk.END, subject)
        
        body_text.delete(1.0, tk.END)
        body_text.insert(tk.END, body)

# Set up the GUI window
root = tk.Tk()
root.title("Email Generator")
root.geometry("600x400")

# Add the input field for the user's name
name_label = tk.Label(root, text="Enter Your Name:")
name_label.pack(pady=10)

name_entry = tk.Entry(root, width=50)
name_entry.pack(pady=5)

# Add a "Browse" button to open the file dialog
browse_button = tk.Button(root, text="Browse for File", command=browse_file)
browse_button.pack(pady=10)

# Add a label and text box to display the subject
subject_label = tk.Label(root, text="Generated Subject:")
subject_label.pack(pady=10)

subject_text = tk.Text(root, height=3, width=70)
subject_text.pack(pady=5)

# Add a label and text box to display the body
body_label = tk.Label(root, text="Generated Body:")
body_label.pack(pady=10)

body_text = tk.Text(root, height=8, width=70)
body_text.pack(pady=5)

# Start the GUI loop
root.mainloop()
