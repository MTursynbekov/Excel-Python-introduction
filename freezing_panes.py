import openpyxl

wb = openpyxl.load_workbook('produceSales.xlsx')
sheet = wb.active

# Freezing Panes.
sheet.freeze_panes = 'A2' # Freeze the rows above A2

# Frozen Pane Examples.
# sheet.freeze_pane = 'A2' # Freeze the row 1.
# sheet.freeze_pane = 'B1' # Freeze the column A.
# sheet.freeze_pane = 'C1' # Freeze the columns A and B.
# sheet.freeze_pane = 'C2' # Freeze the row 1 and columns A and B.
# sheet.freeze_pane = 'A1' or sheet.freeze_pane = None # No frozen panes.

wb.save('freezeExample.xlsx')
