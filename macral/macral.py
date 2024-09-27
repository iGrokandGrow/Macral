import os
from docx import Document
import commonmark
import tkinter as tk
from tkinter import *
from tkinter import filedialog
from tkinter import ttk

#Function to convert docx to markdown
def docx_to_markdown(docx_path):
    doc = Document(docx_path)
    full_text = []
    for para in doc.paragraphs:
        full_text.append(para.text)
    # Convert the plain text to Obsidian readable markdown
    markdown_text = commonmark.commonmark('\n'.join(full_text))
    return markdown_text

#Function to process all docx files in a folder
def process_folder(folder_path):
    for filename in os.listdir(folder_path):
        if filename.endswith('.docx'):
            docx_path = os.path.join(folder_path, filename)
            markdown_text = docx_to_markdown(docx_path)
            markdown_filename = os.path.splitext(filename) [0] + '.md'
            markdown_path = os.path.join(folder_path, markdown_filename)
            
            with open(markdown_path, 'w', encoding='utf-8') as md_file:
                md_file.write(markdown_text)

#Create the main Tkinter window and interface:

root = tk.Tk()
root.title("Docx to Markdown Converter")

frame = ttk.Frame(root, padding ="10")
frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

folder_path = tk.StringVar()

def select_folder():
    folder_selected = filedialog.askdirectory()
    folder_path.set(folder_selected)
    
def start_conversion():
    if folder_path.get():
        process_folder(folder_path.get())
        result_label.config(text="Conversion completed!")
    else:
        result.label.config(text="Please select a folder first.")
select_button = ttk.Button(frame, text="Select Folder",
                           command=select_folder)
select_button.grid(column=0, row=0, padx=5, pady=5)
path_label = ttk.Label(frame, textvariable=folder_path)
path_label.grid(column=1, row=0, padx=5, pady=5)

convert_button = ttk.Button(frame, text="Convert", command=start_conversion)
convert_button.grid(column=0, row=1, padx=5, pady=5)

result_label = ttk.Label(frame, text="")
result_label.grid(column=1, row=1, padx=5, pady=5)

root.mainloop()

