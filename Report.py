# Excel reporting app

# import openpyxl module
import openpyxl

dd = "09"
mm = "02"
yyyy = "2023"

# Give the location of the file
path_E1 = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\E1 " + dd + "." + mm + "." + yyyy + ".xlsx"
path_E2 = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\E2 " + dd + "." + mm + "." + yyyy + ".xlsx"
path_daily = "C:\\Dane\\Python\\Wlasne\\Report\\Raporty Inbet\\" + dd + "." + mm + "." + yyyy + ".xlsx"

# To open the workbook. Workbook object is created
wb_obj = openpyxl.load_workbook(path_daily)

# Get workbook active sheet object from the active attribute
# sheet_obj = wb_obj.active
sheet_obj = wb_obj['E2']    #działa

# Cell objects also have a row, column, and coordinate attributes that provide location information for the cell.

# Note: The first row or column integer is 1, not 0.

# Cell object is created by using sheet object's cell() method.
cell_obj = sheet_obj.cell(row=26, column=3)

# Print value of cell object
# using the value attribute
print(cell_obj.value)
# print(sheet_obj.max_row)
# print(sheet_obj.max_column)

cell_obj.value = "39999"           # było 39
wb_obj.save(path_daily)


