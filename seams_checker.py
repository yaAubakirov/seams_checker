import tkinter as tk
import tkinter.filedialog as dg
import tkinter.scrolledtext as st
import os
import re
import base64
import threading
import openpyxl
import fitz


# this class is used as storage for variables to pass them them between classes
class Storage:
    text = None
    weld_list = None
    ndt_list = None
    temp_drawing_number_list = None
    filename = None


# class which one works with text
class Analyze:
    # this method extracts text from pdf and push it to dictionary
    @classmethod
    def extract_text_from_pdf2(cls, pdf_path):
        with fitz.open(pdf_path) as doc:
            text = {}
            for num, page in enumerate(doc):
                text[num] = page.getText()

            Storage.text = text

    # this class gets weld number, its index and text. Looks for concatenated weld number with ndt class in text
    @classmethod
    def find_in_text(cls, to_find, index, text):
        # list with three types of concat
        ndt_classes = [
            ' ' + str(Storage.ndt_list[index]),
            str(Storage.ndt_list[index]),
            ' ' + str(Storage.ndt_list[index])
        ]
        for ndt in ndt_classes:
            for page in text:
                try:
                    res_search = re.search(to_find + ndt, text[page])
                except:
                    return False

                if res_search:
                    return True


# main class which one runs application interface
class App:
    def __init__(self, master):
        # icon base64 file
        file = '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 '
        icon_data = base64.b64decode(file)
        # the temp file is icon.ico
        temp_file = "icon.ico"
        icon_file = open(temp_file, "wb")
        # Extract the icon
        icon_file.write(icon_data)
        icon_file.close()

        # application interface
        self.master = master
        master.title('Welds checker')
        master.geometry("600x400")
        master.resizable(0, 0)
        master.columnconfigure(2, weight=2)
        master.iconbitmap(temp_file)
        # Delete the temp file
        os.remove(temp_file)

        self.load_pdf = tk.Button(master, text='PDF', height=1, width=10, bd=1, command=self.pdf_load)
        self.load_pdf.grid(row=1, column=2, pady=4, padx=4)

        self.load_excel = tk.Button(master, text='WSL', height=1, width=10, bd=1, command=self.excel_load)
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
        Storage.filename = filename

    def excel_load(self):
        try:
            file = dg.askopenfile(mode='rb', title='Choose WSL report', filetypes=[("Excel files", ".xlsx")])
        except PermissionError:
            self.txt.insert('end', "Close WSL first\n")
            return False
        if file is not None:
            filepath = os.path.abspath(file.name)
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            max_rows = ws.max_row
            temp_welds_list = []
            temp_ndt_list = []
            temp_drawing_number_list = []
            for row in ws.iter_rows(min_row=2, min_col=12, max_col=12, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            temp_welds_list.append(cell.value)
                        else:
                            self.txt.insert('end', "Spaces should be deleted from WSL report\n")
                            return False

            for row in ws.iter_rows(min_row=2, min_col=20, max_col=20, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            temp_ndt_list.append(cell.value)
                        else:
                            self.txt.insert('end', "Spaces should be deleted from WSL report\n")
                            return False

            for row in ws.iter_rows(min_row=2, min_col=4, max_col=4, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            temp_drawing_number_list.append(cell.value)
                        else:
                            self.txt.insert('end', "Spaces should be deleted from WSL report\n")
                            return False
            # checking for cases when drawing numbers are not filled in the end
            if len(temp_welds_list) > len(temp_drawing_number_list):
                a = len(temp_welds_list) - len(temp_drawing_number_list)
                for i in range(a):
                    temp_drawing_number_list.append('Bom-bom-bom')

            if len(temp_welds_list) > len(temp_ndt_list):
                a = len(temp_welds_list) - len(temp_ndt_list)
                for i in range(a):
                    temp_ndt_list.append('XXX')
            # put lists to storage
            Storage.weld_list = temp_welds_list
            Storage.ndt_list = temp_ndt_list
            Storage.temp_drawing_number_list = temp_drawing_number_list
        else:
            self.txt.insert('end', "WSL is not uploaded\n")
            return False
        self.txt.insert('end', "WSL is uploaded\n")

    def analyze(self):
        wrong_welds = []
        self.txt.insert('end', ".............\n")
        text = ""
        if Storage.text is None:
            self.txt.insert('end', "PDF is not loaded\n")
        else:
            text = Storage.text

        if Storage.weld_list:
            for index, weld in enumerate(Storage.weld_list):
                if Analyze.find_in_text(str(weld), index, text):
                    if Storage.temp_drawing_number_list[index] == Storage.filename[:37]:
                        self.refresh()
                        self.txt.insert('end', "{} is OK\n".format(str(weld)))
                    else:
                        self.refresh()
                        self.txt.insert('end',
                                        "Erection drawing number for {} in WSL is not correct\n".format(str(weld)),
                                        'warning')
                        wrong_welds.append(weld)
                else:
                    self.txt.insert('end', "Problem with {}\n".format(str(weld)), 'warning')
                    wrong_welds.append(weld)
        else:
            self.txt.insert('end', "WSL is not loaded\n")

        self.txt.insert('end', ".............\n")
        self.txt.insert('end', "\nTotal count of welds is {}\n".format(len(Storage.weld_list)))
        self.txt.insert('end', "Total count of not found welds is {}\n\n".format(len(wrong_welds)))

        if len(wrong_welds) == 0:
            self.txt.insert('end', ".............\n\n")
            self.txt.insert('end', "{} is OK\n".format(Storage.filename))
        if len(Storage.weld_list) == len(wrong_welds):
            self.txt.insert('end', "Probably spaces are not deleted from WSL\n")

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
