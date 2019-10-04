import xlsxwriter
from itertools import zip_longest
import os
import xlrd
from datetime import date
from datetime import timedelta


def comp_xls(wb1, wb2):
    path_dest = "C:\\Users\\" + os.getlogin() + "\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Comp_" + str(
        date.today()) + ".xlsx"

    wb = xlsxwriter.Workbook(path_dest)

    red = wb.add_format({'font_color': 'red', 'font_size': 10})
    red_date_form = wb.add_format({'num_format': 'dd/mm/yyy', 'font_color': 'red', 'font_size': 10})
    green = wb.add_format({'font_color': 'green', 'font_size': 10})
    orange = wb.add_format({'font_color': 'orange', 'font_size': 10})
    title = wb.add_format({'bold': True, 'font_size': 11})
    regular_text = wb.add_format({'font_size': 10, 'align': 'vcenter'})
    date_form = wb.add_format({'num_format': 'dd/mm/yyy', 'font_size': 10})

    for sheet1, sheet2 in zip_longest(wb1.sheets(), wb2.sheets(), fillvalue=xlrd.sheet.Sheet):
        if wb1.nsheets < wb2.nsheets:
            sheet_c = wb.add_worksheet(sheet2.name)
        else:
            sheet_c = wb.add_worksheet(sheet1.name)
        if sheet1.nrows != sheet2.nrows:
            sheet_c.set_tab_color('red')
            for c in range(max(sheet1.ncols, sheet2.ncols)):
                for r in range(min(sheet1.nrows, sheet2.nrows)):
                    if sheet1.cell_value(r, c) == sheet2.cell_value(r, c):
                        if r == 0:
                            sheet_c.write(r, c, sheet1.cell_value(r, c), title)
                        elif type(sheet1.cell_value(r, c)) == float:
                            sheet_c.write(r, c, sheet2.cell_value(r, c), date_form)
                        else:
                            sheet_c.write(r, c, sheet1.cell_value(r, c), regular_text)
                    else:
                        if type(sheet1.cell_value(r, c)) == float and type(sheet2.cell_value(r, c)) == float:
                            sheet_c.write(r, c, date.fromtimestamp((sheet2.cell_value(r, c)-25568)*86400).strftime('%d/%m/%Y') + " ---> " + date.fromtimestamp((sheet1.cell_value(r, c)-25568)*86400).strftime('%d/%m/%Y'), red)
                        else:
                            sheet_c.write(r, c, sheet2.cell_value(r, c) + " ---> " + sheet1.cell_value(r, c), red)
                for r in range(min(sheet1.nrows, sheet2.nrows), max(sheet1.nrows, sheet2.nrows)):
                    if sheet1.nrows > sheet2.nrows:
                        if type(sheet1.cell_value(r, c)) == float:
                            sheet_c.write(r, c, date.fromtimestamp((sheet1.cell_value(r, c) - 25568) * 86400), red_date_form)
                        else:
                            sheet_c.write(r, c, sheet1.cell_value(r, c), red)
                    else:
                        if type(sheet2.cell_value(r, c)) == float:
                            sheet_c.write(r, c, date.fromtimestamp((sheet2.cell_value(r, c) - 25568) * 86400),
                                          red_date_form)
                        else:
                            sheet_c.write(r, c, sheet2.cell_value(r, c), red)
        else:
            for c in range(max(sheet1.ncols, sheet2.ncols)):
                for r in range(max(sheet1.nrows, sheet2.nrows)):
                    if sheet1.cell_value(r, c) == sheet2.cell_value(r, c):
                        if r == 0:
                            sheet_c.write(r, c, sheet1.cell_value(r, c), title)
                        elif type(sheet1.cell_value(r, c)) == float:
                            sheet_c.write(r, c, sheet2.cell_value(r, c), date_form)
                        else:
                            sheet_c.write(r, c, sheet1.cell_value(r, c), regular_text)
                    else:
                        sheet_c.set_tab_color('red')
                        #if type(sheet1.cell_value(r, c)) == float and type(sheet2.cell_value(r, c)) == float:
                        if c == 11 or c == 12:
                            try:
                                sheet_c.write(r, c, date.fromtimestamp((sheet2.cell_value(r, c)-25568)*86400).strftime('%d/%m/%Y') + " ---> " + date.fromtimestamp((sheet1.cell_value(r, c)-25568)*86400).strftime('%d/%m/%Y'), red)
                            except OSError:
                                sheet_c.write(r, c, "Argument invalide", red)
                        else:
                            sheet_c.write(r, c, sheet2.cell_value(r, c) + " ---> " + sheet1.cell_value(r, c), red)
    wb.close()


def main_comp():
    if date.weekday(date.today()) == 0:
        path1 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_"+ str(
            date.today()) +".xlsx"
        path2 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_"+ str(
            date.today() + timedelta(days=-39)) +".xlsx"

        wb1 = xlrd.open_workbook(path1)
        wb2 = xlrd.open_workbook(path2)
        """i = 0
        while True:
            try:
                wb2 = xlrd.open_workbook(path2)
            except FileNotFoundError:
                i += 1
                path2 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_" + str(
                    date.today() + timedelta(days=-3-i)) + ".xlsx"
            break"""

        comp_xls(wb1, wb2)
    elif 0 < date.weekday(date.today()) < 5:
        path1 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_" + str(
            date.today()) + ".xlsx"
        path2 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_" + str(
            date.today() + timedelta(days=-1)) + ".xlsx"

        wb1 = xlrd.open_workbook(path1)
        wb2 = xlrd.open_workbook(path2)
        """i = 0
        while True:
            try:
                wb2 = xlrd.open_workbook(path2)
            except FileNotFoundError:
                i += 1
                path2 = "C:\\Users\\m.rosillette\\Collectivite Territoriale de Martinique\\SANT - Documents\\DEPLOIEMENT FTTH\\Script\\Résultats\\PRO\\Result_Check_EPRO_" + str(
                    date.today() + timedelta(days=-1 - i)) + ".xlsx"
            break

        comp_xls(wb1, wb2)"""
        comp_xls(wb1, wb2)

# main_comp()