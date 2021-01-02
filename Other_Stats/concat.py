import os
import pandas as pd
os.chdir(r'C:\Users\Kay\Desktop\Data Projects\DraftPick') 

excel_names = []
for i in range(1,11):
    excel_names.append("Modern_Draft_Position_" + str(i) + ".xlsx")
    
excels = [pd.ExcelFile(name) for name in excel_names]

frames = [x.parse(x.sheet_names[0], header=None,index_col=None)for x in excels]
frames[1:] = [df[2:] for df in frames[1:]]
combined = pd.concat(frames)
combined.to_excel("Modern_Draft_Position_All.xlsx", header=False, index=False)