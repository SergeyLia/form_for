import tkinter as tk
import docx
import os
import win32api

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
    show_saved_template()

def show_saved_template():
    saved_doc = docx.Document("filled_template.docx")
    template_text = "\n".join([paragraph.text for paragraph in saved_doc.paragraphs])
    result_label.config(text=template_text)

def print_template():
    file_path = os.path.abspath("filled_template.docx")
    win32api.ShellExecute(0, "print", file_path, None, ".", 0)

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

print_button = tk.Button(root, text="Печать", command=print_template)
print_button.pack()

root.mainloop()