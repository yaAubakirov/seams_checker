import tkinter as tk
import tkinter.filedialog as dg
import tkinter.scrolledtext as st
import os
import re
import threading
import openpyxl
import fitz


class Storage:
    text = None
    weld_list = None


class Analyze:
    @classmethod
    def extract_text_from_pdf2(cls, pdf_path):
        with fitz.open(pdf_path) as doc:
            text = {}
            for num, page in enumerate(doc):
                text[num] = page.getText()

            Storage.text = text

    @classmethod
    def find_in_text(cls, to_find, text):
        ndt_classes = [' A', ' B', ' C', ' D', 'A', 'B', 'C', 'D', ' A', ' B', ' C', ' D']
        for ndt in ndt_classes:
            for page in text:
                try:
                    res_search = re.search(to_find + ndt, text[page])
                except:
                    return False
                if res_search:
                    return True


class App:
    def __init__(self, master):
        self.master = master
        master.title('Welds checker')
        master.geometry("600x400")
        master.resizable(0, 0)
        master.columnconfigure(2, weight=2)

        self.load_pdf = tk.Button(master, text='PDF', height=1, width=10, bd=1, command=self.pdf_load)
        self.load_pdf.grid(row=1, column=2, pady=4, padx=4)

        self.load_excel = tk.Button(master, text='Excel', height=1, width=10, bd=1, command=self.excel_load)
        self.load_excel.grid(row=2, column=2, pady=4, padx=4)

        self.red_button = tk.Button(master, text='Check', height=1, width=10, bd=1, command=self.analyze)
        self.red_button.grid(row=3, column=2, pady=4, padx=4)

        self.txt = st.ScrolledText(master, width=40)
        self.txt.grid(rowspan=5, column=0, row=0, pady=4, padx=4)
        self.txt.tag_config('warning', foreground="red")

        self.copyright = tk.Label(master, text='Metal Yapı Engineering & Construction LLC', fg="#808080")
        self.copyright.place(relx=.6, rely=.95)

    def pdf_load(self):
        file = dg.askopenfile(mode='rb', title='Choose a file', filetypes=[("PDF files", ".pdf")])
        filepath = os.path.abspath(file.name)
        filename = os.path.splitext(os.path.basename(filepath))[0]
        pdf_load = threading.Thread(target=Analyze.extract_text_from_pdf2, args=[filepath])
        pdf_load.daemon = True
        pdf_load.start()
        while pdf_load.is_alive():
            self.loading()
        self.txt.insert('end', "{} is loaded\n".format(filename))

    def excel_load(self):
        file = dg.askopenfile(mode='rb', title='Choose a file', filetypes=[("Excel files", ".xlsx")])
        if file:
            filepath = os.path.abspath(file.name)
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            max_rows = ws.max_row
            temp_list = []
            for row in ws.iter_rows(min_row=2, min_col=12, max_col=12, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            temp_list.append(cell.value)
                        else:
                            self.txt.insert('end', "Spaces should be deleted from WSL report\n")
                            return False

            Storage.weld_list = temp_list
        self.txt.insert('end', "Excel is loaded\n")

    def analyze(self):
        wrong_welds = []
        self.txt.insert('end', "...........\n")
        text = ""
        if Storage.text is None:
            self.txt.insert('end', "PDF is not loaded\n")
        else:
            text = Storage.text

        if Storage.weld_list:
            for weld in Storage.weld_list:
                if Analyze.find_in_text(str(weld), text):
                    self.refresh()
                    self.txt.insert('end', "{} is OK\n".format(str(weld)))
                else:
                    self.txt.insert('end', "Problem with {}\n".format(str(weld)), 'warning')
                    wrong_welds.append(weld)
        else:
            self.txt.insert('end', "Excel is not loaded\n")

        self.txt.insert('end', "\n...........\n")
        self.txt.insert('end', "Total count of welds is {}\n".format(len(Storage.weld_list)))
        self.txt.insert('end', "Total count of not found welds is {}\n".format(len(wrong_welds)))

        if len(wrong_welds) == 0:
            self.txt.insert('end', "...........\n")
            self.txt.insert('end', "Erection drawing is OK\n")

    def loading(self):
        self.txt.insert('end', 'Pdf document processing |')
        self.refresh()
        self.clear_all_text()
        self.txt.insert('end', 'Pdf document processing /')
        self.refresh()
        self.clear_all_text()
        self.txt.insert('end', 'Pdf document processing —')
        self.refresh()
        self.clear_all_text()
        self.txt.insert('end', 'Pdf document processing \\')
        self.refresh()
        self.clear_all_text()

    def refresh(self):
        self.master.update()

    def clear_all_text(self):
        self.txt.delete('1.0', tk.END)


root = tk.Tk()
my_gui = App(root)
root.mainloop()
