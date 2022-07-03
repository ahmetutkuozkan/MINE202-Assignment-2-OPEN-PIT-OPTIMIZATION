import xlsxwriter as excel_writer
Table=excel_writer.Workbook("ECONOMIC BLOCK VALUES(Ore waster drawer with EBV).xlsx")
Sheets=Table.add_worksheet()
cost_for_first_row=8 #$/ton
cost_for_row_increase=0.25#$/ton/bench
unit_ore_processing_cost=27.00#$/ton
Ton_for_waste=(15**3)*2.5
Ton_for_ore=(15**3)*2.1
Revenue=0.04*0.85*1700 #REV = gxTonxP
#revenue1=0
#cost_all=0
for j in range(11):
    for i in range(21):
        if(6<i<14 and 2<j):
            #revenue1+=Revenue
            #print(Revenue,j+1,i+1)
            #print((cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost),j+1,i+1)#Costs
            #cost_all+=(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost)
            Sheets.write(j,i,(Revenue-(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost)))
        else:
            #cost_all+=(cost_for_first_row+(j*cost_for_row_increase))
            Sheets.write(j,i,-(cost_for_first_row+(j*cost_for_row_increase)))
Table.close()
#print(revenue1)
#print(cost_all)