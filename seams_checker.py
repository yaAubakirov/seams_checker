import tkinter as tk
import tkinter.filedialog as dg
import tkinter.scrolledtext as st
import os
import sys
import re
import random
import base64
import threading
import openpyxl
import fitz


# this class is used as storage for variables to pass them between classes
class Storage:
    text = None
    weld_list = None
    ndt_list = None
    first_mark_list = None
    second_mark_list = None
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
                text[num] = page.get_text()

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

    @classmethod
    def weld_without_ndt(cls, weld, text):
        for page in text:
            try:
                res_search = re.search(weld, text[page])
            except:
                return False

            if res_search:
                return True

    @classmethod
    def is_weld_is_plating_grating(cls, weld, index, text):
        if Analyze.weld_without_ndt(weld, text):
            if Storage.first_mark_list[index][:3] == "FLP" \
                    or Storage.second_mark_list[index][:3] == "FLP" \
                    or Storage.first_mark_list[index][:2] == "GR" \
                    or Storage.second_mark_list[index][:2] == "GR":
                return True
        else:
            return False


# main class which one runs application interface
class App:
    def __init__(self, master):
        version = 1.20

        datafile = "my.ico"
        if not hasattr(sys, "frozen"):
            datafile = os.path.join(os.path.dirname(__file__), datafile)
        else:
            datafile = os.path.join(sys.prefix, datafile)

        # application interface
        self.master = master
        master.title('Welds checker (v{})'.format(version))
        master.geometry("600x400")
        master.resizable(0, 0)
        master.columnconfigure(2, weight=2)
        master.iconbitmap(default=datafile)

        self.load_pdf = tk.Button(master, text='PDF', height=1, width=10, bd=1, command=self.pdf_load)
        self.load_pdf.grid(row=1, column=2, pady=4, padx=4)

        self.load_excel = tk.Button(master, text='WSL', height=1, width=10, bd=1, command=self.excel_load)
        self.load_excel.grid(row=2, column=2, pady=4, padx=4)

        self.red_button = tk.Button(master, text='Check', height=1, width=10, bd=1, command=self.analyze)
        self.red_button.grid(row=3, column=2, pady=4, padx=4)

        self.txt = st.ScrolledText(master, width=40)
        self.txt.grid(rowspan=5, column=0, row=0, pady=4, padx=4)
        self.txt.tag_config('warning', foreground="red")
        self.txt.configure(state='disabled')

        self.copyright = tk.Label(master, text='Metal Yapı Engineering & Construction LLC', fg="#808080")
        self.copyright.place(relx=.6, rely=.95)

    def pdf_load(self):
        file = dg.askopenfile(mode='rb', title='Choose a file', filetypes=[("PDF files", ".pdf")])
        if file is not None:
            filepath = os.path.abspath(file.name)
            filename = os.path.splitext(os.path.basename(filepath))[0]
            pdf_load = threading.Thread(target=Analyze.extract_text_from_pdf2, args=[filepath])
            pdf_load.daemon = True
            pdf_load.start()
            while pdf_load.is_alive():
                self.loading()
            self.insert_text('Drawing is uploaded')
            Storage.filename = filename
        else:
            self.insert_text('PDF is not uploaded')
            return False

    def excel_load(self):
        try:
            file = dg.askopenfile(mode='rb', title='Choose WSL report', filetypes=[("Excel files", ".xlsx")])
        except PermissionError:
            self.insert_text('Close WSL first')
            return False
        if file is not None:
            filepath = os.path.abspath(file.name)
            wb = openpyxl.load_workbook(filepath)
            ws = wb.active
            max_rows = ws.max_row
            temp_welds_list = []
            temp_ndt_list = []
            temp_drawing_number_list = []
            first_mark_list = []
            second_mark_list = []
            for row in ws.iter_rows(min_row=2, min_col=12, max_col=12, max_row=max_rows):
                for cell in row:
                    if cell.value:
                        if str(cell.value)[0] != ' ':
                            temp_welds_list.append(cell.value)
                        else:
                            self.insert_text('Spaces should be deleted from WSL report')
                            return False
                    else:
                        temp_welds_list.append('missed weld number')

            for row in ws.iter_rows(min_row=2, min_col=20, max_col=20, max_row=max_rows):
                for cell in row:
                    if cell.value:
                        if str(cell.value)[0] != ' ':
                            temp_ndt_list.append(cell.value)
                        else:
                            self.insert_text('Spaces should be deleted from WSL report')
                            return False
                    else:
                        temp_ndt_list.append('NDT class is missed')

            for row in ws.iter_rows(min_row=2, min_col=4, max_col=4, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            temp_drawing_number_list.append(cell.value)
                        else:
                            self.insert_text('Spaces should be deleted from WSL report')
                            return False
                    else:
                        temp_drawing_number_list.append('Drawing number is missed')

            for row in ws.iter_rows(min_row=2, min_col=6, max_col=6, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            first_mark_list.append(cell.value)
                        else:
                            self.insert_text('Spaces should be deleted from WSL report')
                            return False
                    else:
                        temp_drawing_number_list.append('Drawing number is missed')

            for row in ws.iter_rows(min_row=2, min_col=9, max_col=9, max_row=max_rows):
                for cell in row:
                    if cell.value is not None:
                        if str(cell.value)[0] != ' ':
                            second_mark_list.append(cell.value)
                        else:
                            self.insert_text('Spaces should be deleted from WSL report')
                            return False
                    else:
                        temp_drawing_number_list.append('Drawing number is missed')

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
            Storage.first_mark_list = first_mark_list
            Storage.second_mark_list = second_mark_list
        else:
            self.insert_text('WSL is not uploaded')
            return False
        self.insert_text('WSL is uploaded')

    def analyze(self):
        wrong_welds = []
        typical_welds = []
        self.insert_text('.............')
        text = ""
        if Storage.text is None:
            self.insert_text('PDF is not loaded')
        else:
            text = Storage.text

        if Storage.weld_list:
            for index, weld in enumerate(Storage.weld_list):
                if Analyze.find_in_text(str(weld), index, text):
                    if Storage.temp_drawing_number_list[index] == Storage.filename[:37]:
                        self.weld_text_insert(weld)
                    else:
                        self.refresh()
                        self.txt.insert('end',
                                        "Erection drawing number for {} in WSL is not correct\n".format(str(weld)),
                                        'warning')
                        wrong_welds.append(weld)
                        self.txt.yview('end')
                elif Analyze.is_weld_is_plating_grating(str(weld), index, text):
                    self.typical_weld_text_insert(weld)
                    typical_welds.append(weld)
                else:
                    self.problem_weld_text_insert(weld)
                    wrong_welds.append(weld)
        else:
            self.insert_text('WSL is not loaded')

        self.insert_text('.............')
        final_result = "\nTotal count of welds is {}".format(len(Storage.weld_list))
        self.insert_text(final_result)
        if len(typical_welds) > 0:
            final_result_for_typical_welds = "Total count of typical welds is {}".format(len(typical_welds))
            self.insert_text(final_result_for_typical_welds)
        final_result_for_wrong_welds = "Total count of problem welds is {}".format(len(wrong_welds))
        self.insert_text(final_result_for_wrong_welds, 2)

        if len(wrong_welds) == 0:
            self.insert_text('.............', 2)
            self.txt.configure(state='normal')
            self.txt.insert('end', "{} ✔\n".format(Storage.filename[:37]), 'name')
            self.txt.tag_config('name', foreground='green')
            self.txt.yview('end')
            self.txt.configure(state='disabled')
        if len(Storage.weld_list) == len(wrong_welds):
            self.insert_text('Probably spaces are not deleted from WSL')
        if len(wrong_welds) > 0:
            self.insert_text('.............', 2)
            self.txt.configure(state='normal')
            self.txt.insert('end', "{} ✘\n".format(Storage.filename[:37]), 'warning')
            self.txt.yview('end')
            self.txt.configure(state='disabled')
        self.txt.yview('end')

    def loading(self):
        hashtags = random.randint(5, 35)
        spaces = 40 - hashtags - 4
        percentage = int(hashtags * 2.5)
        self.txt.configure(state='normal')
        self.txt.insert('end', 'Pdf pages loading. Please wait...\n')
        self.txt.insert('end', '#' * hashtags + ' ' * spaces + '{}%'.format(percentage))
        self.refresh()
        self.clear_all_text()
        self.txt.configure(state='disabled')

    def insert_text(self, text, number_of_n=1):
        self.txt.configure(state='normal')
        self.txt.insert('end', "{}{}".format(text, '\n' * number_of_n))
        self.txt.yview('end')
        self.txt.configure(state='disabled')

    def weld_text_insert(self, weld_number):
        self.txt.configure(state='normal')
        self.refresh()
        self.txt.insert('end', "{} is OK\n".format(str(weld_number)))
        self.txt.yview('end')
        self.txt.configure(state='disabled')

    def problem_weld_text_insert(self, weld_number):
        self.txt.configure(state='normal')
        self.refresh()
        self.txt.insert('end', "Problem with {}\n".format(str(weld_number)), 'warning')
        self.txt.yview('end')
        self.txt.configure(state='disabled')

    def typical_weld_text_insert(self, weld_number):
        self.txt.configure(state='normal')
        self.refresh()
        self.txt.insert('end', "{} is typical\n".format(str(weld_number)), 'attention')
        self.txt.tag_config('attention', foreground='#FFC000')
        self.txt.yview('end')
        self.txt.configure(state='disabled')

    def refresh(self):
        self.master.update()

    def clear_all_text(self):
        self.txt.delete('1.0', tk.END)


root = tk.Tk()
my_gui = App(root)
root.mainloop()
