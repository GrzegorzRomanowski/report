# Excel reporting app

import tkinter as tk
import openpyxl
from openpyxl.styles import PatternFill, Border, Side, Font


# Adjusting style
# color
yellow_fill = PatternFill(start_color='FFFFFF00', end_color='FFFFFF00', fill_type='solid')
zlecenie_fill = PatternFill(start_color='DCDCDC', end_color='DCDCDC', fill_type='solid')
title_fill = PatternFill(start_color='F0E68C', end_color='F0E68C', fill_type='solid')
summary_fill = PatternFill(start_color='FFC000', end_color='FFC000', fill_type='solid')
pink_fill = PatternFill(start_color='BDD7EE', end_color='BDD7EE', fill_type='solid')
# border
edge = Side(border_style='thin', color='000000')
edge2 = Side(border_style='thick', color='000000')
border = Border(top=edge, bottom=edge, left=edge, right=edge)
border2 = Border(top=edge2, bottom=edge2, left=edge2, right=edge2)
# font
font_bold = Font(bold=True)


# Welcome message box
# ask_window = tk.Tk()
# ask_window.title("inBet  raport")
# ask_window.geometry("250x115")
#
# ask_label = tk.Label(ask_window, text="Za jaki dzień zrobić raport?")
# ask_label.pack(pady=5, padx=5)
#
# ask_frame = tk.Frame(ask_window)
# ask_frame.pack(pady=5, padx=5)
#
# ask_day = tk.Entry(ask_frame, width=6)
# ask_day.grid(pady=2, padx=2, row=0, column=0)
# ask_day.insert(0, "dd")
#
# ask_month = tk.Entry(ask_frame, width=6)
# ask_month.grid(pady=2, padx=2, row=0, column=1)
# ask_month.insert(0, "mm")
#
# ask_year = tk.Entry(ask_frame, width=10)
# ask_year.grid(pady=2, padx=2, row=0, column=2)
# ask_year.insert(0, "yyyy")
# def button():
#     global dd, mm, yyyy
#     dd = ask_day.get()
#     mm = ask_month.get()
#     yyyy = ask_year.get()
#     ask_window.destroy()
#
# ask_button = tk.Button(ask_window, text="Zatwierdź", command=button, width=25, height=2)
# ask_button.pack(pady=5, padx=5)
#
# ask_window.mainloop()

# Warning message box
def warning(w):
    print(w)


# Hardcoded data and variables
# Date
dd = "09"
mm = "02"
yyyy = "2023"
safe = 1
# Columns names
title_tuple = ("Element", "Typ", "Powierzchnia", "Klasa betonu", "Gatunek stali")
title_index = (3, 4, 6, 7, 9)

# Locations of the files
path_E1 = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\E1 " + dd + "." + mm + "." + yyyy + ".xlsx"
path_E2 = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\E2 " + dd + "." + mm + "." + yyyy + ".xlsx"
path_daily = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + ".xlsx"
pow_do_raportu = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Produkcja płyt wg projektów - " + yyyy + " - powierzchnia do raportu.xlsx"


# Load workbooks and sheets
# To open the workbook. Workbook object is created.
wb_E1 = openpyxl.load_workbook(path_E1)
wb_E2 = openpyxl.load_workbook(path_E2)
wb_daily = openpyxl.load_workbook(path_daily)
wb_pow_do_raportu = openpyxl.load_workbook(pow_do_raportu)
# Get workbook active sheet object from the active attribute or sheet name.
sheet_E1 = wb_E1.active
sheet_E2 = wb_E2.active
sheet_daily_E1 = wb_daily['E1']
sheet_daily_E2 = wb_daily['E2']
sheet_pow_do_raportu = wb_pow_do_raportu.active
# row 1 from "sheet_pow_do_raportu"
row_1_pow_do_raportu = sheet_pow_do_raportu[1]
# amount of rows in sheets
E1_max_row = sheet_E1.max_row
E2_max_row = sheet_E2.max_row


# create and fill in E1_list of projects
project_E1_list = []
z = 0
for i in range (1, E1_max_row):
    if (sheet_E1.cell(column=5, row=i).value) == None:
        None
    else:
        project_E1_list.append([sheet_E1.cell(column=5, row=i).value, i])
        if z == 0:
            None
        else:
            project_E1_list[z-1].append(project_E1_list[z][1]-project_E1_list[z-1][1]-6)
        z += 1
project_E1_list[-1].append(E1_max_row-project_E1_list[-1][1]-7)

# create and fill in E2_list of projects
project_E2_list = []
z2 = 0
for i in range (1, E2_max_row):
    if (sheet_E2.cell(column=5, row=i).value) == None:
        None
    else:
        project_E2_list.append([sheet_E2.cell(column=5, row=i).value, i])
        if z2 == 0:
            None
        else:
            project_E2_list[z2-1].append(project_E2_list[z2][1]-project_E2_list[z2-1][1]-6)
        z2 += 1
project_E2_list[-1].append(E2_max_row-project_E2_list[-1][1]-7)

print(project_E1_list)
print(project_E2_list)


# Making a daily report
# copy from E1 report and E2 report to daily report
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
    tt = 0
    for t in title_tuple:
        sheet_daily_E1.cell(row=i[1]+2, column=title_index[tt]).value = t
        tt += 1
    for row in sheet_daily_E1.iter_rows(min_row=i[1]+2, max_row=i[1]+2, min_col=1, max_col=9):
        for cell in row:
            cell.fill = title_fill
    # Summary slabs
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).value = i[2]
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).border = border2
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=3).fill = summary_fill

    # Main data table
    for j in row_1_pow_do_raportu:
        if j.value == i[0]:
            jj = j.column
    y = 0
    while i[2] > y:
        num = int(sheet_E1.cell(row=i[1]+y+3, column=3).value)
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
        if (sheet_pow_do_raportu.cell(row=num+8, column=jj).value) != None:
            if (sheet_pow_do_raportu.cell(row=num + 8, column=jj).fill.start_color.index) != 'FFFFFF00':
                sheet_daily_E1.cell(row=i[1]+y+3, column=6).value = sheet_pow_do_raportu.cell(row=num+8, column=jj).value
                sheet_pow_do_raportu.cell(row=num+8, column=jj).fill = yellow_fill
            else:
                war = "Próbujesz wpisać do raportów płytę,\nktóra już wcześniej była wyprodukowana/zaraportowana:\nProjekt:  " + str(i[0]) + "\nNumer elementu:  " + str(num) + "\nSkrypt się zamknie bez zapisywania żadnych zmian.\nZweryfikuj błąd i uruchom skrypt ponownie."
                warning(war)
                safe = 0
        else:
            for jjj in range(1, 10):
                if (sheet_pow_do_raportu.cell(row=num+8, column=jj+jjj).value) != None:
                    if (sheet_pow_do_raportu.cell(row=num+8, column=jj+jjj).fill.start_color.index) != 'FFFFFF00':
                        print("niestandardowa płyta", num)
                        sheet_daily_E1.cell(row=i[1]+y+3, column=6).value = sheet_pow_do_raportu.cell(row=num+8, column=jj+jjj).value
                        sheet_pow_do_raportu.cell(row=num+8, column=jj+jjj).fill = yellow_fill
                        sheet_daily_E1.cell(row=i[1]+y+3, column=11).value = sheet_pow_do_raportu.cell(row=num+8, column=jj+jjj).value
                        sheet_daily_E1.cell(row=i[1]+y+3, column=10).value = sheet_pow_do_raportu.cell(row=5, column=jj+jjj).value
                        sheet_daily_E1.cell(row=i[1]+y+3, column=10).fill = pink_fill
                        sheet_daily_E1.cell(row=i[1]+y+3, column=11).fill = pink_fill
                        break
                    else:
                        war = "Próbujesz wpisać do raportów płytę,\nktóra już wcześniej była wyprodukowana/zaraportowana:\nProjekt:  " + str(i[0]) + "\nNumer elementu:  " + str(num) + "\nSkrypt się zamknie bez zapisywania żadnych zmian.\nZweryfikuj błąd i uruchom skrypt ponownie."
                        warning(war)
                        safe = 0
        y += 1

    # Summary area
    sum_list = []
    k = 0
    for kkk in range(0, i[2]):
        sum_list.append(i[1]+3+k)
        k += 1
    stri = "=0"
    for kk in sum_list:
        stri += "+F" + str(kk)
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).value = stri
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).fill = summary_fill
    sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).border = border2
    # sheet_daily_E1.cell(row=i[1]+i[2]+4, column=6).value = "=SUMA(F" + str(i[1]+3) + ":F" + str(i[1]+i[2]+2) + ")"

# Heading
sheet_daily_E1['H8'] = dd + "." + mm + "." + yyyy
sheet_daily_E1['H10'] = "E1"
sheet_daily_E2['H8'] = dd + "." + mm + "." + yyyy
sheet_daily_E2['H10'] = "E2"

# Reducing row height
small_rows = (4, 5, 6, 7, 9, 11)
for i in small_rows:
    sheet_daily_E1.row_dimensions[i].height = 1

# Saving files
if safe == 1:
    wb_daily.save("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + "GRZ.xlsx")
    wb_pow_do_raportu.save("C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\Pow_do_raportu_GRZ.xlsx")

#dostęp do adresu komórki
# print(cell_obj2.row, cell_obj2.column)

# pobranie koloru
# ccc = sheet_pow_do_raportu['C11'].fill.start_color.index #Yellow Color
# print(ccc, type(ccc))
# wypełnienie kolorem
# yellow_fill = PatternFill(start_color=ccc, end_color=ccc, fill_type='solid')
# sheet_daily_E1['H10'].fill = yellow_fill

# obramowanie wszystkich komórek z wartościami (bez poscalanych niestety)
# for row in sheet_daily_E1.iter_rows(min_row=1, min_col=1, max_col=9, max_row=E1_max_row):
#     for cell in row:
#         if cell.value != None:
#             cell.border = border