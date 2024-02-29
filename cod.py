import os
import tkinter as tk
from tkinter import filedialog
from docx import Document
from io import BytesIO

def extract_images_from_docx(docx_file, output_folder):
    doc = Document(docx_file)

    images = {}
    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            images[rel.target_part.partname] = image_data

    for idx, (partname, image_data) in enumerate(sorted(images.items())):
        
        image_name = f"extracted_image_{idx + 1}.png"
            
        image_path = os.path.join(output_folder, image_name)
            
        with open(image_path, "wb") as img_file:
            img_file.write(image_data)

def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    entry_path.delete(0, tk.END)
    entry_path.insert(tk.END, file_path)

def browse_folder():
    folder_path = filedialog.askdirectory()
    entry_folder.delete(0, tk.END)
    entry_folder.insert(tk.END, folder_path)

def extract_images():
    docx_file_path = entry_path.get()
    output_folder_path = entry_folder.get()

    if os.path.isfile(docx_file_path) and os.path.isdir(output_folder_path):
        extract_images_from_docx(docx_file_path, output_folder_path)
        result_label.config(text="Изображения успешно извлечены и сохранены.")
    else:
        result_label.config(text="Пожалуйста, выберите правильный файл и папку.")

root = tk.Tk()
root.title("Извлечение изображений из .docx")


label_path = tk.Label(root, text="Выберите .docx файл:")
label_path.grid(row=0, column=0, padx=10, pady=10, sticky=tk.W)

entry_path = tk.Entry(root, width=40)
entry_path.grid(row=0, column=1, padx=10, pady=10)

button_browse_file = tk.Button(root, text="Обзор", command=browse_file)
button_browse_file.grid(row=0, column=2, padx=10, pady=10)

label_folder = tk.Label(root, text="Выберите папку для сохранения изображений:")
label_folder.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)

entry_folder = tk.Entry(root, width=40)
entry_folder.grid(row=1, column=1, padx=10, pady=10)

button_browse_folder = tk.Button(root, text="Обзор", command=browse_folder)
button_browse_folder.grid(row=1, column=2, padx=10, pady=10)

button_extract = tk.Button(root, text="Извлечь изображения", command=extract_images)
button_extract.grid(row=2, column=0, columnspan=3, pady=20)

result_label = tk.Label(root, text="")
result_label.grid(row=3, column=0, columnspan=3)

root.mainloop()