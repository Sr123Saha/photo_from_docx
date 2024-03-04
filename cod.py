import os
import tkinter as tk
from tkinter import filedialog
from tkinter import font
from docx import Document

# Функция для извлечения изображений из файла .docx и сохранения их в указанную папку
def extract_images_from_docx(docx_file, output_folder):
    doc = Document(docx_file)

    images = {}
    image_idx = 1  # Индекс изображений

    for rel in doc.part.rels.values():
        if "image" in rel.reltype:
            image_data = rel.target_part.blob
            image_name = f"product{image_idx}.png"
            image_path = os.path.join(output_folder, image_name)
            with open(image_path, "wb") as img_file:
                img_file.write(image_data)
            images[rel.target_part.partname] = image_path
            image_idx += 1

    return images
    # Проходим по всем связям в документе
    for rel in doc.part.rels.values():
        # Если тип связи - изображение
        if "image" in rel.reltype:
            # Получаем данные изображения
            image_data = rel.target_part.blob
            # Сохраняем данные в словарь с использованием имен связей
            images[rel.target_part.partname] = image_data

    # Сохраняем изображения в указанную папку
    for idx, (partname, image_data) in enumerate(sorted(images.items())):
        image_name = f"product{idx + 1}.png"
        image_path = os.path.join(output_folder, image_name)

        with open(image_path, "wb") as img_file:
            img_file.write(image_data)

# Функция для выбора файла .docx
def browse_file():
    file_path = filedialog.askopenfilename(filetypes=[("Word files", "*.docx")])
    entry_path.config(state=tk.NORMAL)  # Разрешаем ввод
    entry_path.delete(0, tk.END)
    entry_path.insert(tk.END, file_path)
    entry_path.config(state="readonly")  # Запрещаем ввод

# Функция для выбора папки
def browse_folder():
    folder_path = filedialog.askdirectory()
    entry_folder.config(state=tk.NORMAL)  # Разрешаем ввод
    entry_folder.delete(0, tk.END)
    entry_folder.insert(tk.END, folder_path)
    entry_folder.config(state="readonly")  # Запрещаем ввод

# Функция для извлечения изображений после выбора файла и папки
def extract_images():
    docx_file_path = entry_path.get()
    output_folder_path = entry_folder.get()

    if os.path.isfile(docx_file_path) and os.path.isdir(output_folder_path):
        # Извлекаем изображения и выводим успешное сообщение
        extract_images_from_docx(docx_file_path, output_folder_path)
        result_label.config(text="Изображения успешно извлечены и сохранены.")
    else:
        # Выводим сообщение об ошибке, если файл или папка не выбраны
        result_label.config(text="Пожалуйста, выберите правильный файл и папку.")

# Создание основного окна Tkinter
root = tk.Tk()
root.geometry("700x150")
root.title("Извлечение изображений из .docx", )
root.configure(bg="#0ECCA6")

# Настройка шрифтов
bold_font = font.Font(weight='bold')
custom_font = font.Font(size=11)

# Label и Entry для выбора файла .docx
label_path = tk.Label(root, text="Выберите .docx файл:", bg="#0ECCA6",font=custom_font)
label_path.grid(row=0, column=0, padx=10, pady=1, sticky=tk.W)
entry_path = tk.Entry(root, width=40, bg="#F0F0F0", state="readonly")  # Запрет ввода
entry_path.grid(row=0, column=1, padx=10, pady=10)
button_browse_file = tk.Button(root, text="Обзор", command=browse_file, bg='#10CEE3')
button_browse_file.grid(row=0, column=2, padx=10, pady=10)

# Label и Entry для выбора папки
label_folder = tk.Label(root, text="Выберите папку для сохранения изображений:", bg="#0ECCA6",font=custom_font)
label_folder.grid(row=1, column=0, padx=10, pady=10, sticky=tk.W)
entry_folder = tk.Entry(root, width=40 ,bg="#F0F0F0", state="readonly")  # Запрет ввода
entry_folder.grid(row=1, column=1, padx=10, pady=10)
button_browse_folder = tk.Button(root, text="Обзор", command=browse_folder, bg='#10CEE3')
button_browse_folder.grid(row=1, column=2, padx=10, pady=10)

# Button для извлечения изображений
button_extract = tk.Button(root, text="Извлечь изображения", command=extract_images,bg='#10CEE3',font=bold_font)
button_extract.grid(row=2, column=0, columnspan=3, pady=20)

# Label для вывода результата
result_label = tk.Label(root, text="", bg='#0ECCA6')
result_label.grid(row=3, column=0, columnspan=3)

# Запуск главного цикла Tkinter
root.mainloop()
