import xlsxwriter as excel_writer
Table=excel_writer.Workbook("Table(Ore waster drawer).xlsx")
Sheets=Table.add_worksheet()

for j in range(11):
    for i in range(21):
        if(6<i<14 and 2<j):
            Sheets.write(j,i,"O")
        else:
            Sheets.write(j,i,"W")
Table.close()
