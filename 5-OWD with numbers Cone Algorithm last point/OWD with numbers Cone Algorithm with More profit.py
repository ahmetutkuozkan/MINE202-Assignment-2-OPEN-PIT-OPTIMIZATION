import xlsxwriter as excel_writer
Table=excel_writer.Workbook("ECONOMIC BLOCK VALUES(OWD with numbers Cone Algorithm with More profit).xlsx")
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
            Sheets.write(j,i,(Revenue-(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost))*Ton_for_ore)
            Sheets_one_by_one=((Revenue-(cost_for_first_row+j*cost_for_row_increase+unit_ore_processing_cost))*Ton_for_ore)
        else:
            #cost_all+=(cost_for_first_row+(j*cost_for_row_increase))*Ton_for_waste
            Sheets.write(j,i,-(cost_for_first_row+(j*cost_for_row_increase))*Ton_for_waste)
            Sheets_one_by_one=(-(cost_for_first_row+(j*cost_for_row_increase))*Ton_for_waste)
        Sheets_in_python_row.append(Sheets_one_by_one)
    Sheets_in_python.append(Sheets_in_python_row)
Table.close()
#print(Sheets_in_python)
#print(revenue1)
#print(cost_all)
control_index=[]
save_index=[]
save=0
control=0
for row in range(len(Sheets_in_python)):
    for column in range(len(Sheets_in_python[0])):
        control=0
        control_index=[]
        try:
            if(Sheets_in_python[row][column] >0):

                for row_excavate in range(row+1):
                    for column_excavate in range(column-row+row_excavate,column+row-row_excavate+1):
                        if(column_excavate<0):
                            control-=100000000000
                        #print(Sheets_in_python[row_excavate][column_excavate],row_excavate,column_excavate,end=" ")
                        control+=Sheets_in_python[row_excavate][column_excavate]
                        control_index.append((row_excavate+1,column_excavate+1))

                    #print(" ")

        except IndexError:
            control=0
            control_index=[]
        if(control>save):
            save=control
            #print(control,control_index)
            save_index=control_index
            control_index=[]

find_last_index=[]
find_last_index_list=[]
find_last_index_list.append(save_index[len(save_index)-1])
find_last_index=save_index[len(save_index)-1]

#Last digged place position
#print(save,save_index)
for last in range(len(save_index)-2,0,-1):
    if(find_last_index[0]==save_index[last][0]):
        find_last_index_list.append(save_index[last])
print(find_last_index_list)









