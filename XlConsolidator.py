from tkinter import *
from tkinter import filedialog
import os
import datetime
import xlrd, datetime, xlwt
from xlwt import easyxf
from xlutils.copy import copy as xl_copy


HEADER = 2
EMPTY_CELL = ""
key_delimiter = "##"

DATE_cno = 0
CONSIGNEE_cno = 1
VOUCHER_cno = 2
GSTIN_cno = 3
GROSS_TOTAL_cno = 4
VAL_18_cno = 5
VAL_28_cno = 6
SERVICE_CHARGE_cno = 7
O_SGST_9_cno = 8
O_CGST_9_cno = 9
O_SGST_14_cno = 10
O_CGST_14_cno = 11
O_IGST_28_cno = 12
O_IGST_18_cno = 13
ROUND_OFF_cno = 14


DATE_txt = "Date"
CONSIGNEE_txt = "Consignee/Buyer"
VOUCHER_txt = "Voucher No."
GSTIN_txt = "GSTIN/UIN"
GROSS_TOTAL_txt = "Gross Total"
VAL_18_txt = "Sale 18%"
VAL_28_txt = "Sale 28%"
ROUND_OFF_txt = "R/O"
SERVICE_CHARGE_txt = "Serv.Ch."
CGST_txt = "CGST"
SGST_txt = "SGST"
IGST_txt = "IGST"
CASH_SALES = "Cash Sales"
SUNDRY_DEBTORS = "Sundry Debtors"
text_cno_mapping = [(GROSS_TOTAL_cno, GROSS_TOTAL_txt), (VAL_18_cno, VAL_18_txt), (VAL_28_cno, VAL_28_txt), (SERVICE_CHARGE_cno, SERVICE_CHARGE_txt)]
numeric_cols = [GROSS_TOTAL_cno, VAL_18_cno, VAL_28_cno, SERVICE_CHARGE_cno, O_CGST_9_cno, O_CGST_14_cno, O_SGST_9_cno,
                O_SGST_14_cno, O_IGST_18_cno, O_IGST_28_cno, ROUND_OFF_cno]
final = {}
index_tracker = {}

class XlConsolidator(Tk):
    def __init__(self):
        Tk.__init__(self)
        self.winfo_toplevel().title("Excel Consolidator")
        self.winfo_toplevel().resizable(0, 0)
        self.winfo_toplevel().geometry("+450+300")

        self._frame = None
        self.switch_frame(StartPage)
        menu_frame = Frame(self)
        menu_frame.grid(row=0, column=0, padx=1, pady=1)

        XlConsolidator.read_wb = None
        excel_file_path = StringVar()
        excel_file_path_label = Label(self, text="Excel file Path", anchor=NE)
        excel_file_path_label.grid(row=0, column=0, sticky=W)
        excel_file_path_entry = Entry(self, textvariable=excel_file_path,
                                     bd=1, width=30, relief=SUNKEN, state='readonly')
        excel_file_path_entry.xview_moveto(0.5)
        excel_file_path_entry.grid(row=0, column=1)
        Button(self, text="Browse", fg="black", bd=0, bg="#fff", cursor="hand2",
               command=lambda: XlConsolidator.browse_btn(excel_file_path)) \
            .grid(row=0, column=2, padx=1, pady=1, sticky=SE)
        Button(self, text="Execute", fg="black", bd=0, bg="#fff", cursor="hand2",
               command=lambda: XlConsolidator.consolidate(excel_file_path)) \
            .grid(row=1, column=2, padx=1, pady=1, sticky=SE)

        status_frame = Frame(self)
        status_frame.grid(row=2, column=0, columnspan=2, padx=1, pady=1, sticky=W)

        status_label = Label(status_frame, text="Status:", bd=1, width=5)
        status_label.grid(row=0, column=0, padx=5, pady=1)

        self.status_var = StringVar()
        self.status_var.set("Application Initiated")
        statusbar = Entry(status_frame, textvariable=self.status_var, bd=1, width=42, relief=SUNKEN,
                          state='readonly')
        scroll = Scrollbar(self, orient='horizontal', command=statusbar.xview)
        statusbar.config(xscrollcommand=scroll.set)

        statusbar.grid(row=0, column=1, padx=1, pady=1, sticky=W)

    def switch_frame(self, frame_class):
        """Destroys current frame and replaces it with a new one."""
        new_frame = frame_class(self)
        if self._frame is not None:
            self._frame.destroy()
        self._frame = new_frame
        self._frame.grid(row=0, column=1, padx=1, pady=1)

    @staticmethod
    def browse_btn(var):
        file_path = filedialog.askopenfilename()
        if file_path:
            var.set(file_path)

    @staticmethod
    def consolidate(var):
        input_file = var.get()
        XlConsolidator.read_wb = xlrd.open_workbook(input_file)
        read_sheet = XlConsolidator.read_wb.sheet_by_index(0)

        for i in range(HEADER + 1, read_sheet.nrows):
            if read_sheet.cell_value(i, 0) is not EMPTY_CELL:
                XlConsolidator.add_to_final(read_sheet.row(i))
            else:
                break
        write_book = xl_copy(XlConsolidator.read_wb)
        # write_book = xlwt.Workbook()
        write_sheet = write_book.add_sheet(sheetname="reformated")
        row_index = 0
        write_sheet.write(row_index, 0, read_sheet.cell_value(0, 0))
        row_index += 1
        write_sheet.write(row_index, 0, read_sheet.cell_value(1, 0))

        # write header
        row_index += 1
        write_sheet.write(row_index, 0, DATE_txt)
        write_sheet.write(row_index, 1, CONSIGNEE_txt)
        write_sheet.write(row_index, 2, VOUCHER_txt)
        col_index = 3
        for tuple in text_cno_mapping:
            write_sheet.write(row_index, col_index, tuple[1])
            col_index += 1
        write_sheet.write(row_index, col_index, CGST_txt)
        write_sheet.write(row_index, col_index + 1, SGST_txt)
        write_sheet.write(row_index, col_index + 2, IGST_txt)
        write_sheet.write(row_index, col_index + 3, ROUND_OFF_txt)

        # write data
        style_blue = xlwt.easyxf(
            'pattern: pattern solid, fore_colour light_blue;' 'font: name arial narrow, height 180, colour white, bold True;')
        style_white = easyxf('pattern: pattern solid, fore_colour white')
        old_date = datetime.date.today()
        odd = 1
        row_index += 1
        for key, value in final.items():
            date = key.split(key_delimiter)[1]
            if old_date != date:
                odd += 1
                old_date = date
            consignee = key.split(key_delimiter)[2]
            write_sheet.write(row_index, 0, date, style_blue if odd % 2 == 0 else style_white)
            write_sheet.write(row_index, 1, consignee, style_blue if odd % 2 == 0 else style_white)
            write_sheet.write(row_index, 2, value[VOUCHER_txt], style_blue if odd % 2 == 0 else style_white)
            col_index = 3
            for tuple in text_cno_mapping:
                write_sheet.write(row_index, col_index, value[tuple[1]], style_blue if odd % 2 == 0 else style_white)
                col_index += 1

            write_sheet.write(row_index, col_index, value[CGST_txt], style_blue if odd % 2 == 0 else style_white)
            write_sheet.write(row_index, col_index + 1, value[SGST_txt], style_blue if odd % 2 == 0 else style_white)
            write_sheet.write(row_index, col_index + 2, value[IGST_txt], style_blue if odd % 2 == 0 else style_white)
            write_sheet.write(row_index, col_index + 3, value[ROUND_OFF_txt],
                              style_blue if odd % 2 == 0 else style_white)
            row_index += 1
        write_sheet.write(row_index, 1, "Total:")
        for i in range(2, 10):
            write_sheet.write(row_index, i, xlwt.Formula(f"SUM(${XlConsolidator.get_column(i)}$1:"
                                                         f"${XlConsolidator.get_column(i)}${row_index})"))

        
        import os
        filename, file_extension = os.path.splitext(input_file)
        write_book.save(filename + "-edited.xls")
        app.quit()

    @staticmethod
    # write formula
    def get_column(colno):
        switcher = {
            2: "C",
            3: "D",
            4: "E",
            5: "F",
            6: "G",
            7: "H",
            8: "I",
            9: "J"
        }
        return switcher.get(colno)

    def log(self, msg):
        self.status_var.set(msg)

    @staticmethod
    def get_surrogate_key(date, consignee, gst):
        global index_tracker

        index = 0
        if consignee != CASH_SALES and gst is EMPTY_CELL:
            consignee = SUNDRY_DEBTORS
        if consignee != CASH_SALES and consignee != SUNDRY_DEBTORS:
            temp_key = str(date) + key_delimiter + consignee
            index = index_tracker[temp_key] + 1 if temp_key in index_tracker else index
            index_tracker[str(date) + key_delimiter + consignee] = index
        return str(index) + key_delimiter + str(date) + key_delimiter + consignee

    @staticmethod
    def get_cell_val(row, cno):
        cell_val = row[cno].value
        if (type(cell_val) is not float or int) and cell_val is EMPTY_CELL and cno in numeric_cols:
            cell_val = 0.0
        return cell_val

    @staticmethod
    def get_CGST(row):
        return round(XlConsolidator.get_cell_val(row, O_CGST_9_cno), 2) + \
               round(XlConsolidator.get_cell_val(row, O_CGST_14_cno), 2)

    @staticmethod
    def get_SGST(row):
        return round(XlConsolidator.get_cell_val(row, O_SGST_9_cno), 2) + \
               round(XlConsolidator.get_cell_val(row, O_SGST_14_cno), 2)

    @staticmethod
    def get_IGST(row):
        return round(XlConsolidator.get_cell_val(row, O_IGST_18_cno), 2) + \
               round(XlConsolidator.get_cell_val(row, O_IGST_28_cno), 2)

    @staticmethod
    def add_to_final(row):
        global final
        raw_date = row[DATE_cno].value
        consignee = row[CONSIGNEE_cno].value
        gst = row[GSTIN_cno].value
        date_as_datetime = datetime.datetime(*xlrd.xldate_as_tuple(raw_date, XlConsolidator.read_wb.datemode))
        date = date_as_datetime.date()
        surrogate_key = XlConsolidator.get_surrogate_key(date, consignee, gst)

        if surrogate_key not in final:
            temp = {}
            for tuple in text_cno_mapping:
                try:
                    cell_val = XlConsolidator.get_cell_val(row, tuple[0])
                except Exception as e:
                    print(row[tuple[0]].value)
                    print(tuple[0])
                    print(surrogate_key)
                    print(row)
                    print(e)
                    pass
                cell_val = round(cell_val, 2)
                temp[tuple[1]] = cell_val
            temp[VOUCHER_txt] = row[VOUCHER_cno].value
            temp[GSTIN_txt] = row[GSTIN_cno].value
            temp[CGST_txt] = XlConsolidator.get_CGST(row)
            temp[SGST_txt] = XlConsolidator.get_SGST(row)
            temp[IGST_txt] = XlConsolidator.get_IGST(row)
            temp[ROUND_OFF_txt] = XlConsolidator.get_cell_val(row, ROUND_OFF_cno)
            final[surrogate_key] = temp
        else:
            temp = final[surrogate_key]
            for tuple in text_cno_mapping:
                cell_val = XlConsolidator.get_cell_val(row, tuple[0])
                try:
                    temp[tuple[1]] = temp[tuple[1]] + cell_val
                except:
                    print(surrogate_key)
                    print(final[surrogate_key])
                    print(tuple)
                    print(temp[tuple[1]])
                    print(cell_val)
                    print(row)
                    pass
            temp[CGST_txt] = temp[CGST_txt] + XlConsolidator.get_CGST(row)
            temp[SGST_txt] = temp[SGST_txt] + XlConsolidator.get_SGST(row)
            temp[IGST_txt] = temp[IGST_txt] + XlConsolidator.get_IGST(row)
            temp[ROUND_OFF_txt] = temp[ROUND_OFF_txt] + XlConsolidator.get_cell_val(row, ROUND_OFF_cno)
            final[surrogate_key] = temp


class StartPage(Frame):
    def __init__(self, master):
        Frame.__init__(self, master)
        Label(self, text="This is the start page").grid(row=0, column=0, padx=1, pady=1)



if __name__ == "__main__":
    # global app
    app = XlConsolidator()
    app.mainloop()
