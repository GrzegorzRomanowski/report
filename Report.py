# Excel reporting script

import tkinter as tk
import sys
import os
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font
from pycel import ExcelCompiler


# Adjusting style
# color
yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
zlecenie_fill = PatternFill(start_color='DCDCDC', end_color='DCDCDC', fill_type='solid')
title_fill = PatternFill(start_color='F0E68C', end_color='F0E68C', fill_type='solid')
summary_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
pink_fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
green_fill = PatternFill(start_color='548235', end_color='548235', fill_type='solid')
# border
edge = Side(border_style='thin', color='000000')
edge2 = Side(border_style='thick', color='000000')
border = Border(top=edge, bottom=edge, left=edge, right=edge)
border2 = Border(top=edge2, bottom=edge2, left=edge2, right=edge2)
# font
font_bold = Font(bold=True)


# Hardcoded data and variables
# Date
# dd = "09"
# mm = "02"
# yyyy = "2023"
# ebawe = "1"
safe_to_save = 1
area_found = 0
additional_sums_written = 0
# Columns names
title_tuple = ("Element", "Typ", "Powierzchnia", "Klasa betonu", "Gatunek stali")
title_index = (3, 4, 6, 7, 9)


# Welcome message box
ask_window = tk.Tk()
ask_window.title("inBet  Report")
ask_window.geometry("250x195")

ask_label = tk.Label(ask_window, text="Za jaki dzień zrobić raport?\nZ której linii produkcyjnej?")
ask_label.pack(pady=5, padx=5)

ask_frame = tk.Frame(ask_window)
ask_frame.pack(pady=5, padx=5)

ask_day = tk.Entry(ask_frame, width=6)
ask_day.grid(pady=2, padx=2, row=0, column=0)
ask_day.insert(0, "dd")

ask_month = tk.Entry(ask_frame, width=6)
ask_month.grid(pady=2, padx=2, row=0, column=1)
ask_month.insert(0, "mm")

ask_year = tk.Entry(ask_frame, width=10)
ask_year.grid(pady=2, padx=2, row=0, column=2)
ask_year.insert(0, "yyyy")

ebawe_label = tk.Label(ask_frame, text="EBAWE: ")
ebawe_label.grid(pady=18, padx=2, row=1, columnspan=2, column=0)

ask_ebawe = tk.Entry(ask_frame, width=10)
ask_ebawe.grid(pady=18, padx=2, row=1, column=2)
ask_ebawe.insert(0, "1")


def button():
    global dd, mm, yyyy, ebawe
    dd = ask_day.get()
    mm = ask_month.get()
    yyyy = ask_year.get()
    ebawe = ask_ebawe.get()
    ask_window.destroy()


ask_button = tk.Button(ask_window, text="Zatwierdź", command=button, width=25, height=2)
ask_button.pack(pady=5, padx=5)

ask_window.mainloop()


# Warning message box
def warning(w):
    warning_window = tk.Tk()
    warning_window.title("Warning !")

    warning_label = tk.Label(warning_window, text=w)
    warning_label.pack(pady=20, padx=25)
    print(w)

    def warning_button():
        global safe_to_save
        safe_to_save = 0
        warning_window.destroy()
        sys.exit()
    warning_button = tk.Button(warning_window, text="OK", command=warning_button, width=25, height=2)
    warning_button.pack(pady=10, padx=15)

    warning_window.mainloop()


# Locations of the files
path_E1 = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\E" + ebawe + " " + dd + "." + mm + "." + yyyy + ".xlsx"


if os.path.isfile("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + "_GR.xlsx"):
    path_daily = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + "_GR.xlsx"
else:
    path_daily = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Szablon.xlsx"


if os.path.isfile("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Produkcja płyt wg projektów - " + yyyy + " - powierzchnia do raportu_GR.xlsx"):
    pow_do_raportu = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Produkcja płyt wg projektów - " + yyyy + " - powierzchnia do raportu_GR.xlsx"
else:
    pow_do_raportu = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Produkcja płyt wg projektów - " + yyyy + " - powierzchnia do raportu.xlsx"


if os.path.isfile("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + mm + "." + yyyy + " - zestawienie miesięczne, tutaj sumy z raportów_GR.xlsx"):
    path_month = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + mm + "." + yyyy + " - zestawienie miesięczne, tutaj sumy z raportów_GR.xlsx"
else:
    path_month = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + mm + "." + yyyy + " - zestawienie miesięczne, tutaj sumy z raportów.xlsx"

# Load workbooks and sheets
try:
    wb_E1 = openpyxl.load_workbook(path_E1)
except (FileNotFoundError, NameError):
    war = """Nie znaleziono pliku z raportem EBAWE.

Skrypt zamknie się bez zapisywania żadnych zmian.

Sprawdź czy raport jest stworzony.
Zweryfikuj jego nazwę, rozszerzenie, lokalizację itp.
i uruchom skrypt ponownie."""

    warning(war)
wb_daily = openpyxl.load_workbook(path_daily)
wb_pow_do_raportu = openpyxl.load_workbook(pow_do_raportu)
wb_month = openpyxl.load_workbook(path_month)
# Get workbook active sheet object from the active attribute or sheet name.
sheet_E1 = wb_E1.active
sheet_daily_E1 = wb_daily['E'+ebawe]
sheet_pow_do_raportu = wb_pow_do_raportu.active
sheet_month_E1 = wb_month['E'+ebawe]
# row 1 variables
row_1_pow_do_raportu = sheet_pow_do_raportu[1]
row_1_month_E1 = sheet_month_E1[1]
# amount of rows in sheets
E1_max_row = sheet_E1.max_row
pow_do_raportu_max_row = sheet_pow_do_raportu.max_row
pow_do_raportu_max_col = sheet_pow_do_raportu.max_column


# Create and fill in E1_list of projects
project_E1_list = []
z = 0
for i in range(1, E1_max_row):
    if sheet_E1.cell(column=5, row=i).value is not None:
        project_E1_list.append([sheet_E1.cell(column=5, row=i).value, i])
        if z == 0:
            None    # intentional omission of the loop operation in the first iteration
        else:
            project_E1_list[z-1].append(project_E1_list[z][1]-project_E1_list[z-1][1]-6)
        z += 1
project_E1_list[-1].append(E1_max_row-project_E1_list[-1][1]-7)


# Making a daily report - copy from E1 report to daily report

# For loop for every project in E1 report (project_E1_list)
for i in project_E1_list:
    # Zlecenie
    sheet_daily_E1.cell(row=i[1], column=1).value = "Zlecenie:"
    sheet_daily_E1.cell(row=i[1], column=5).value = i[0]
    sheet_daily_E1.cell(row=i[1], column=1).font = font_bold
    sheet_daily_E1.cell(row=i[1], column=5).font = font_bold
    for row in sheet_daily_E1.iter_rows(min_row=i[1], max_row=i[1], min_col=1, max_col=9):
        for cell in row:
            cell.fill = zlecenie_fill
    # Row dimensions
    sheet_daily_E1.row_dimensions[i[1]+1].height = 4
    sheet_daily_E1.row_dimensions[i[1]+i[2]+3].height = 4
    # Columns titles
    t = 0
    for title in title_tuple:
        sheet_daily_E1.cell(row=i[1]+2, column=title_index[t]).value = title
        t += 1
    for row in sheet_daily_E1.iter_rows(min_row=i[1]+2, max_row=i[1]+2, min_col=1, max_col=9):
        for cell in row:
            cell.fill = title_fill
    # Sum of elements
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).value = i[2]
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).border = border2
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).fill = summary_fill
    # Finding proper column index in "pow_do_raportu"
    for j in row_1_pow_do_raportu:
        if j.value == i[0]:
            jj = j.column
    try:
        jj += 0
    except NameError:
        war = "Nie znaleziono odpowiedniego projektu\nw Excelu z powierzchniami brutto.\n\n" + str(i[0])\
              + """\n\nSkrypt zamknie się bez zapisywania żadnych zmian.

Sprawdź, czy numery projektów są identyczne
we wszystkich wejściowych arkuszach excel
i uruchom skrypt ponownie."""
        warning(war)
    # Project description
    description = sheet_pow_do_raportu.cell(row=2, column=jj).value.replace("\n", " - ")
    sheet_daily_E1.cell(row=i[1], column=10).value = description
    i.append(description)
    # Main data table
    # While loop for every prefab element in the project
    y = 0
    while i[2] > y:
        area_found = 0
        try:
            num = int(sheet_E1.cell(row=i[1]+y+3, column=3).value)
        except:
            war = "Nr elementu:  " + str(sheet_E1.cell(row=i[1]+y+3, column=3).value) + "\nz projektu:  " + str(i[0])\
                  + """\n\nnie jest liczbą.
Skrypt zamknie się bez zapisywania żadnych zmian.
Sprawdź numery elementów i uruchom skrypt ponownie."""
            warning(war)
        sheet_daily_E1.cell(row=i[1]+y+3, column=3).value = sheet_E1.cell(row=i[1]+y+3, column=3).value
        sheet_daily_E1.cell(row=i[1]+y+3, column=3).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=4).value = sheet_E1.cell(row=i[1]+y+3, column=4).value
        sheet_daily_E1.cell(row=i[1]+y+3, column=4).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=5).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=7).value = sheet_E1.cell(row=i[1]+y+3, column=9).value
        sheet_daily_E1.cell(row=i[1]+y+3, column=7).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=8).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=9).value = sheet_E1.cell(row=i[1]+y+3, column=11).value
        sheet_daily_E1.cell(row=i[1]+y+3, column=9).border = border
        sheet_daily_E1.cell(row=i[1]+y+3, column=6).border = border
        if sheet_pow_do_raportu.cell(row=num+8, column=jj).value is not None:
            if sheet_pow_do_raportu.cell(row=num + 8, column=jj).fill.start_color.index != 'FFFFFF00':
                sheet_daily_E1.cell(row=i[1]+y+3, column=6).value = sheet_pow_do_raportu.cell(row=num+8, column=jj).value
                sheet_pow_do_raportu.cell(row=num+8, column=jj).fill = yellow_fill
            else:
                war = "Próbujesz wpisać do raportów płytę,\nktóra już wcześniej była zaraportowana:\n\nProjekt:  "\
                      + str(i[0]) + "\nNumer elementu:  " + str(num)\
                      + """\n\nSkrypt zamknie się bez zapisywania żadnych zmian.

Zweryfikuj błąd i uruchom skrypt ponownie."""
                warning(war)
        else:
            # Looking for the area in the next 10 columns on the right
            for index_increase in range(1, 10):
                if sheet_pow_do_raportu.cell(row=num+8, column=jj+index_increase).value is not None:
                    if sheet_pow_do_raportu.cell(row=num+8, column=jj+index_increase).fill.start_color.index != 'FFFFFF00':
                        sheet_daily_E1.cell(row=i[1]+y+3, column=6).value = sheet_pow_do_raportu.cell(row=num+8, column=jj+index_increase).value
                        sheet_pow_do_raportu.cell(row=num+8, column=jj+index_increase).fill = yellow_fill
                        sheet_daily_E1.cell(row=i[1]+y+3, column=11).value = sheet_pow_do_raportu.cell(row=num+8, column=jj+index_increase).value
                        sheet_daily_E1.cell(row=i[1]+y+3, column=10).value = sheet_pow_do_raportu.cell(row=5, column=jj+index_increase).value
                        sheet_daily_E1.cell(row=i[1]+y+3, column=10).fill = pink_fill
                        sheet_daily_E1.cell(row=i[1]+y+3, column=11).fill = pink_fill
                        area_found = 1
                        break
                    else:
                        war = "Próbujesz wpisać do raportów płytę,\nktóra już wcześniej była zaraportowana:\n\nProjekt:  "\
                              + str(i[0]) + "\nNumer elementu:  " + str(num)\
                              + """\n\nSkrypt zamknie się bez zapisywania żadnych zmian.

Zweryfikuj błąd i uruchom skrypt ponownie."""
                        warning(war)
            if area_found != 1:
                war = "Nie można pobrać powierzchni\nelementu nr:  " + str(num) + "\nz projektu:  " + str(i[0])\
                      + """\n\nSkrypt zamknie się bez zapisywania żadnych zmian.

Sprawdź poprawność numeru tego elementu\ni uruchom skrypt ponownie."""
                warning(war)
        y += 1
    # Creating a sum formula
    sum_list = []
    for element in range(0, i[2]):
        sum_list.append(i[1]+3+element)
    formula = "=0"
    for row_index_to_sum in sum_list:
        formula += "+F" + str(row_index_to_sum)
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).value = formula
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).fill = summary_fill
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).border = border2
    # Formula below looks better, but doesn't work ("apparent intersection error")
    # sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).value = "=SUMA(F" + str(i[1]+3) + ":F" + str(i[1]+i[2]+2) + ")"

    # Additional Sums
    add_sum_dict = {}
    # For loop for initializing keys
    for col in sheet_daily_E1.iter_cols(min_col=10, max_col=10, min_row=i[1]+3, max_row=i[1]+i[2]+2):
        for cell in col:
            if cell.value is not None:
                add_sum_dict[cell.value] = 0
    # For loop for set values to keys
    for col in sheet_daily_E1.iter_cols(min_col=10, max_col=10, min_row=i[1]+3, max_row=i[1]+i[2]+2):
        for cell in col:
            if cell.value is not None:
                add_sum_dict[cell.value] += sheet_daily_E1.cell(row=cell.row, column=cell.column+1).value
    i.append(add_sum_dict)
    for key in add_sum_dict.keys():
        additional_sums_written = 0
        for col in sheet_daily_E1.iter_cols(min_col=10, max_col=10, min_row=i[1]+3, max_row=i[1]+i[2]+2):
            for cell in col:
                if cell.value == key and additional_sums_written == 0:
                    sheet_daily_E1.cell(row=cell.row, column=cell.column+2).value = add_sum_dict[key]
                    sheet_daily_E1.cell(row=cell.row, column=cell.column+2).fill = summary_fill
                    sheet_daily_E1.cell(row=cell.row, column=cell.column).fill = summary_fill
                    additional_sums_written = 1

# Heading
sheet_daily_E1['H8'] = dd + "." + mm + "." + yyyy
sheet_daily_E1['H10'] = "E" + str(ebawe)

# Reducing row height
small_rows = (4, 5, 6, 7, 9, 11)
for i in small_rows:
    sheet_daily_E1.row_dimensions[i].height = 1


# Painting green "pow_do_raportu"
for col in sheet_pow_do_raportu.iter_cols(min_row=9, min_col=2, max_col=pow_do_raportu_max_col, max_row=pow_do_raportu_max_row):
    proj_done = 1
    for cell in col:
        if cell.value is not None and cell.fill.start_color.index != 'FFFFFF00':
            proj_done = 0
            break
    if proj_done == 1:
        sheet_pow_do_raportu.cell(row=4, column=col[4].column).fill = green_fill
        sheet_pow_do_raportu.cell(row=5, column=col[5].column).fill = green_fill
for proj in project_E1_list:
    for proj2 in row_1_pow_do_raportu:
        if proj[0] == proj2.value:
            list_of_proj_to_paint = [proj2]
            proj2_col = proj2.column
            for u in range(1, 11):
                if sheet_pow_do_raportu.cell(row=1, column=proj2_col+u).value is None:
                    list_of_proj_to_paint.append(sheet_pow_do_raportu.cell(row=1, column=proj2_col+u))
                else:
                    break
            len_to_paint = len(list_of_proj_to_paint)
            paint_or_not = 1
            for u in range(0, len_to_paint):
                if sheet_pow_do_raportu.cell(row=4, column=list_of_proj_to_paint[u].column).fill.start_color.index != '00548235':
                    paint_or_not = 0
            if paint_or_not == 1:
                for proj3 in row_1_month_E1:
                    if proj[0] == proj3.value:
                        proj3_col = proj3.column
                for u in range(0, len_to_paint):
                    sheet_pow_do_raportu.cell(row=1, column=proj2_col+u).fill = green_fill
                    sheet_pow_do_raportu.cell(row=2, column=proj2_col+u).fill = green_fill
                    sheet_pow_do_raportu.cell(row=3, column=proj2_col+u).fill = green_fill
                    # Painting yellow "month report"
                    try:
                        sheet_month_E1.cell(row=1, column=proj3_col+u).fill = yellow_fill
                        sheet_month_E1.cell(row=2, column=proj3_col+u).fill = yellow_fill
                        sheet_month_E1.cell(row=3, column=proj3_col+u).fill = yellow_fill
                        sheet_month_E1.cell(row=4, column=proj3_col+u).fill = yellow_fill
                    except NameError:
                        war = "Nie znaleziono odpowiedniego projektu\nw Excelu z miesięcznym raportem:\n\n" + str(
                            proj[0]) + """\n\nSkrypt zamknie się bez zapisywania żadnych zmian.

Sprawdź, czy projekt znajduje się w tabeli
z miesięcznym raportem i czy jest poprawnie wpisany.
Nastepnie uruchom skrypt ponownie."""
                        warning(war)
# Saving files
if safe_to_save == 1:
    wb_daily.save("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + "_GR.xlsx")
    wb_pow_do_raportu.save("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Produkcja płyt wg projektów - " + yyyy + " - powierzchnia do raportu_GR.xlsx")


# Filling month report
# sheet_month_E1.insert_cols(15) - NIE UŻYWAĆ, BO PSUJE CAŁĄ TABELE EXCELA !!!
excel = ExcelCompiler(filename="C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + "_GR.xlsx")
for i in project_E1_list:
    col_index = "test"
    proj = i[0]
    for cell in row_1_month_E1:
        if cell.value == proj:
            col_index = cell.column
    try:
        col_index += 0
    except (NameError, TypeError):
        war = "Nie znaleziono odpowiedniego projektu\nw Excelu z miesięcznym raportem:\n\n" + str(proj)\
              + """\n\nSkrypt zamknie się bez zapisywania zmian w zestawieniu miesięcznym.
Raport dzienny i tabela z powierzchniami brutto
zostały przeprocesowane prawidłowo.

Sprawdź, czy projekt znajduje się w tabeli
z miesięcznym raportem i czy jest poprawnie wpisany.
Nastepnie uruchom skrypt ponownie."""
        warning(war)
    row_index_to_evaluate = i[1]+i[2]+4
    evaluated_value = excel.evaluate('E' + str(ebawe) + '!F' + str(row_index_to_evaluate))
    print(evaluated_value)
    for dict_value in i[4].values():
        evaluated_value -= dict_value
        print(evaluated_value)
    sheet_month_E1.cell(row=int(dd)+5, column=col_index).value = evaluated_value
    len_dict = len(i[4])
    for el in range(len_dict):
        sheet_month_E1.cell(row=int(dd)+5, column=col_index+el+1).value = i[4][sheet_month_E1.cell(row=2, column=col_index+el+1).value]

# Saving month report
if safe_to_save == 1:
    wb_month.save("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + mm + "." + yyyy + " - zestawienie miesięczne, tutaj sumy z raportów_GR.xlsx")

    warning("JUŻ  :)")
