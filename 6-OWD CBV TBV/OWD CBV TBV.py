import xlsxwriter as excel_writer
Table=excel_writer.Workbook("ECONOMIC BLOCK VALUES(OWD CBV TBV).xlsx")
Sheets=Table.add_worksheet()
cost_for_first_row=8 #$/ton
cost_for_row_increase=0.25#$/ton/bench
unit_ore_processing_cost=27.00#$/ton
Ton_for_waste=(15**3)*2.5
Ton_for_ore=(15**3)*2.1
Revenue=0.04*0.85*1700 #REV = gxTonxP
#revenue1=0
#cost_all=0
Sheets_one_by_one=0

Sheets_in_python=[]
for j in range(11):
    Sheets_in_python_row=[]
    for i in range(21):
        if(6<i<14 and 2<j):
            #revenue1+=Revenue
            #cost_all+=(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost)*Ton_for_ore
            #Sheets.write(j,i,(Revenue-(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost))*Ton_for_ore)
            Sheets_one_by_one=((Revenue-(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost)))
        else:
            #cost_all+=(cost_for_first_row+(j*cost_for_row_increase))*Ton_for_waste
            #Sheets.write(j,i,-(cost_for_first_row+(j*cost_for_row_increase))*Ton_for_waste)
            Sheets_one_by_one=(-(cost_for_first_row+(j*cost_for_row_increase)))
        Sheets_in_python_row.append(Sheets_one_by_one)
    Sheets_in_python.append(Sheets_in_python_row)

#print(Sheets_in_python)
#print(revenue1)
#print(cost_all)
EBV=[]
CBV=[]
TBV=[]
EBV=Sheets_in_python
for j in range(11):
    for i in range(21):
        Sheets.write(3*j+1,i,EBV[j][i])
CBV=EBV
for row in range(len(EBV)):
    for column in range(len(EBV[0])):
        if(row>0):
            CBV[row][column]+=CBV[row-1][column]
for j in range(11):
    for i in range(21):
        Sheets.write(3*j+2,i,CBV[j][i])
TBV=CBV
for column in range(len(CBV[0])):
    for row in range(len(CBV)):
        if(column>0):
            if(row==0):
                TBV_first=TBV[row][column-1]
                TBV_second=TBV[row+1][column-1]
                TBV_third=0
                if(TBV_first>TBV_second and TBV_first>TBV_third):
                    TBV[row][column]+=TBV_first
                    #print(TBV_first,TBV[row][column])
                elif(TBV_second>TBV_first and TBV_second>TBV_third):
                    TBV[row][column]+=TBV_second
                    #print(TBV_second,TBV[row][column])
                elif(TBV_third>TBV_first and TBV_third>TBV_second):
                    TBV[row][column]+=TBV_third
                    #print(TBV_third,TBV[row][column])
            elif(0<row<len(CBV)-1):
                TBV_first=TBV[row-1][column-1]
                TBV_second=TBV[row][column-1]
                TBV_third=TBV[row+1][column-1]
                if(TBV_first>TBV_second and TBV_first>TBV_third):
                    TBV[row][column]+=TBV_first
                elif(TBV_second>TBV_first and TBV_second>TBV_third):
                    TBV[row][column]+=TBV_second
                elif(TBV_third>TBV_first and TBV_third>TBV_second):
                    TBV[row][column]+=TBV_third
            elif(row==(len(CBV)-1)):
                TBV_second=TBV[row-1][column-1]
                TBV_third=TBV[row][column-1]
                if(TBV_second>TBV_third):
                    TBV[row][column]+=TBV_second
                elif(TBV_third>TBV_second):
                    TBV[row][column]+=TBV_third



for j in range(11):
    for i in range(21):
        #Sheets.write(j+3*(j+1),i,EBV[j][i])
        #Sheets.write(j+4*j,i,CBV[j][i])
        Sheets.write(3*j+3,i,TBV[j][i])
Table.close()









