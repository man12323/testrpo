import pandas as pd
import openpyxl

#dates=['14']
dates=['1','2','3','4','5','6','7','8','9','10','11','12','13','14','15','16','17','18','19','20','21','22','23','24','25','26','27','28','29','30']
file1 = open('FTY_JULY_DCT.xlsx')
print(file1.name)
print(type(file1.name))
print(file1.mode)
print(type(file1.mode))

for j in range(0,len(dates)):
    workbook=pd.read_excel('FTY_JULY_DCT.xlsx',sheet_name = dates[j])
    wb=workbook.iloc[34:76,0:18]
    #dt_date=workbook.iloc[10,7]
    #print(type(wb))
    #dt_date1=pd.DataFrame(dt_date)
    #print(dt_date1)
    print(wb)
    wb.rename(columns = {'Unnamed: 1' : 'Product'}, inplace = True)
    df = pd.DataFrame(wb)

    df=df.dropna(subset=['Product'])
 
    #print(df)
    with pd.ExcelWriter('New_File_DCT.xlsx',mode ='a',engine="openpyxl",if_sheet_exists='overlay') as writer:  
        df.to_excel(writer,startrow = writer.sheets['Sheet1'].max_row,header=False, index = False,sheet_name='Sheet1')
        #dt_date.to_excel(writer,startrow = writer.sheets['Sheet1'].max_row,column=1,header=False, index = False,sheet_name='Sheet1')
    j=j+1


