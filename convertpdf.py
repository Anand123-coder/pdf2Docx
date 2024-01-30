import tkinter as tk
from tkinter.filedialog import askopenfile
from tkinter import messagebox
from pdf2docx import Converter

win = tk.Tk()
win.title('Pdf2Docx')

def convertToDocx(pdf_file):
    # Generate the output DOCX file name by replacing the '.pdf' extension with '.docx'
    docx_file = pdf_file.replace('.pdf', '.docx')

    # Convert PDF to DOCX
    cv = Converter(pdf_file)
    cv.convert(docx_file, start=0, end=None)
    cv.close()

    return docx_file

def openFile():
    file = askopenfile(filetypes=[('PDF files', '*.pdf')])
    
    if file:
        docx_file = convertToDocx(file.name)
        messagebox.showinfo('Done', f'File "{file.name}" successfully converted to "{docx_file}"')

label = tk.Label(win, text='Choose PDF File:')
label.grid(row=0, column=0)

button = tk.Button(win, text="----Select----", width=30, command=openFile)
button.grid(row=0, column=1)

win.mainloop()
