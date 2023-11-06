import tkinter as tk
import docx
import os
import subprocess

def submit():
    name = entry_name.get()
    patronymic = entry_patronymic.get()
    surname = entry_surname.get()
    phone = entry_phone.get()
    doc = docx.Document("template1.docx")
    for paragraph in doc.paragraphs:
        paragraph.text = paragraph.text.replace("{{name}}", name)
        paragraph.text = paragraph.text.replace("{{patronymic}}", patronymic)
        paragraph.text = paragraph.text.replace("{{surname}}", surname)
        paragraph.text = paragraph.text.replace("{{phone}}", phone)
    doc.save("filled_template.docx")
    open_saved_template()

def open_saved_template():
    file_path = os.path.abspath("filled_template.docx")
    subprocess.Popen([file_path], shell=True)

root = tk.Tk()


label_name = tk.Label(root, text="Имя:")
label_name.pack()

entry_name = tk.Entry(root)
entry_name.pack()

label_patronymic = tk.Label(root, text="Отчество:")
label_patronymic.pack()

entry_patronymic = tk.Entry(root)
entry_patronymic.pack()

label_surname = tk.Label(root, text="Фамилия:")
label_surname.pack()

entry_surname = tk.Entry(root)
entry_surname.pack()

label_phone = tk.Label(root, text="Телефон:")
label_phone.pack()

entry_phone = tk.Entry(root)
entry_phone.pack()

button = tk.Button(root, text="Ввод", command=submit)
button.pack()

result_label = tk.Label(root, text="")
result_label.pack()


root.mainloop()