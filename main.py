from tkinter import *
from tkinter import filedialog
import pandas as pd
from pandastable import Table
import sys
import os
import traceback

class csvxlsx_convertor:

    def __init__(self, root):
        self.root = root
        self.root.title("Csv Excel Convertor")
        self.main_label = Label(self.root, text="Choose a file to convert")
        self.options_frame = Frame(self.root, height=200, width=300)

        self.table_frame = Frame(root, height=200, width=300)

        self.button_explore_source = Button(self.root, text="Browse Files", command=self.browse_source)

        self.label_file_explorer = Label(self.root, text="Your file Path", width=100, height=4, fg="blue")

        self.teminal = Text(root, height=15, bg="#000000", fg="#10e34c")


        self.button_explore_source.pack()

        self.label_file_explorer.pack()
        self.main_label.pack()
        self.options_frame.pack()
        self.teminal.pack(side='bottom', pady=10)


        self.source_file_path = ""
        self.source_file_sheet_name = ""

        self.dest_file_path = ""
        self.df = None

        self.is_headers = IntVar()
        self.num_of_lines_split = None

        self.export_to_format = None

        sys.stdout = Redirect(self.teminal)



    def browse_source(self):
        try:
            self.source_file_path = filedialog.askopenfilename(initialdir="/", title="Select a File", filetypes=(("Excel files", ".xlsx .xls .csv"),  ("all files", "*.*")))
            # Change label contents
            self.label_file_explorer.configure(text="File Opened: " + self.source_file_path)
            file_path, file_extension = os.path.splitext(self.source_file_path)
            button_restart = Button(root, text="Restart", command=self.restart, bg="#FF0000", fg="#ebebeb")
            button_restart.pack(side="right", pady=10, padx=10)
            self.button_explore_source.destroy()
            # frame.pack()
            if file_extension == ".xlsx" or file_extension == ".xls":
                xl = pd.ExcelFile(self.source_file_path)
                self.main_label.configure(text="Now Choose a Sheet")
                for idx, btn in enumerate(xl.sheet_names):
                    sheet_btn = Button(self.options_frame, text=btn, command=lambda idx=idx: self.choose_dest(xl.sheet_names[idx]), bg="#000000", fg="#ebebeb")
                    sheet_btn.pack(side="right", padx=10)
            else:
                self.choose_dest(self, "")
        except Exception as e:
            print(traceback.format_exc())



    def choose_dest(self, btn):
        try:
            self.clean_frame()
            self.source_file_sheet_name = btn
            self.df = pd.read_csv(self.source_file_path) if btn == "" else pd.read_excel(self.source_file_path, sheet_name=btn)
            self.display_sheet()
            button_explore2 = Button(self.options_frame,
                                     text="choose dest",
                                     command=self.browse_dest)

            button_explore2.pack()
        except Exception as e:
            print(traceback.format_exc())


    def browse_dest(self):
        try:
            dest_folder_path = filedialog.askdirectory(initialdir="/", title="Select a Folder")
            ext = ".xlsx" if self.source_file_sheet_name == "" else ".csv"
            fn_ext = os.path.basename(self.source_file_path)
            fn = fn_ext.rsplit('.', 1)[0]

            # here add file name option
            self.dest_file_path = dest_folder_path + "/" + fn + "-converted" + ext
            self.main_label.configure(text="the file has " + str(len(self.df.index)) + " rows, would you like to split it?")

            self.clean_frame()

            button_split = Button(self.options_frame, text="yes", command=self.split_df)
            button_no_split = Button(self.options_frame, text="no", command=lambda: self.export( self.df, self.dest_file_path, True))
            button_split.pack(side="right", padx=10)
            button_no_split.pack(side="right", padx=10)
        except Exception as e:
            print(traceback.format_exc())


    def split_df(self):
        try:
            self.clean_frame()
            self.num_of_lines_split = Entry(self.options_frame)
            self.num_of_lines_split.pack(side="right", padx=10)

            R1 = Radiobutton(self.options_frame, text="With headers", variable=self.is_headers, value=1)
            R1.pack(side="right", padx=10)

            R2 = Radiobutton(self.options_frame, text="No Headers", variable=self.is_headers, value=2)
            R2.pack(side="right", padx=10)

            export_btn = Button(self.options_frame, text="Export", command=self.pre_export_split)
            export_btn.pack(side="left", padx=10)
        except Exception as e:
            print(traceback.format_exc())


    def pre_export_split(self):
        try:
            num_lines = self.num_of_lines_split .get()
            if (num_lines.isdigit() == False or int(num_lines) > len(self.df.index)):
                self.main_label.configure(text="Split number not valid", fg="#FF0000")
                return
            if (self.is_headers.get() < 1):
                self.main_label.configure(text="please check heades option", fg="#FF0000")
                return
            is_header_export = True if self.is_headers.get() == 1 else False

            chunks = self.split_dataframe(int(num_lines))
            for idx, dfr in enumerate(chunks):
                file_path, file_extension = os.path.splitext(self.dest_file_path)
                new_path = file_path + "-converted-" + str(idx + 1) + file_extension
                self.export( dfr, new_path, is_header_export)
        except Exception as e:
            print(traceback.format_exc())



    def export(self, df, path, isHeader):
        try:
            file_path, file_extension = os.path.splitext(path)
            self.export_to_format = file_extension
            if (self.export_to_format == ".csv"):
                df.to_csv(path, index=None, header=isHeader, encoding='utf-8-sig')
            else:
                df.to_excel(path, index=None, header=isHeader, encoding='utf-8-sig')
            myLabel3 = Label(root, text="done")
            self.clean_frame()
            myLabel3.pack()
        except Exception as e:
            print(traceback.format_exc())



    def split_dataframe(self, chunk_size):
        try:
            chunks = list()
            num_chunks = len(self.df) // chunk_size + 1
            for i in range(num_chunks):
                chunks.append(self.df[i * chunk_size:(i + 1) * chunk_size])
            return chunks
        except Exception as e:
            print(traceback.format_exc())



    def display_sheet(self):
        try:
            if (len(self.df) == 0):
                self.main_label.configure(text="No records', 'No records")
            else:
                pass
            self.table_frame.pack(fill=BOTH, expand=1, side='bottom')
            display_table = Table(self.table_frame, dataframe=self.df, read_only=True)
            display_table.show()
        except Exception as e:
            print(traceback.format_exc())



    def clean_frame(self):
        try:
            for widget in self.options_frame.winfo_children():
                widget.destroy()
            self.options_frame.pack_forget()
            self.options_frame.pack()
        except Exception as e:
            print(traceback.format_exc())


    def restart(self):

        python = sys.executable
        os.execl(python, python, *sys.argv)

        


class Redirect():

    def __init__(self, widget):
        self.widget = widget

    def write(self, text):
        self.widget.insert('end', text)
        # self.widget.see('end') # autoscroll

    def flush(self):
        pass


if __name__ == '__main__':
    root = Tk()
    inst = csvxlsx_convertor(root)
    root.geometry()
    root.mainloop()
