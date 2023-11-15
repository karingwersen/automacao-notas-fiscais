import openpyxl

if __name__ == '__main__':

    path = "C:\\Users\\karin\\Downloads\\PLANILHA NF'S 2018.xlsx"
    planilha = openpyxl.load_workbook(path)
    sheet = planilha.active

    max_row = sheet.max_row
    for i in range(1, max_row + 1):
        cell = sheet.cell(row = i, column = 1)
        print(cell.value)

