# -*- coding: utf-8 -*-
"""
Created on Tue Aug 18 00:36:31 2020

@author: PRANAVSAI
"""

fb = load_workbook(mainlink+"\\Final Aegon.xlsx")

fs = fb.get_sheet_by_name("2")

index=0
a=[]
for ind in range(len(excel)):
    exc = excel[index]
    
    lb = load_workbook(mainlink+"/"+exc)
    
    
    cpos=colind(exc,2)
    
    ls= lb.get_sheet_by_name(str(7))  
    dum1,col=wordfinder("Grand Total (A+B)",1,ls.max_row)
    row1,dum=wordfinder("Individual agents",1,ls.max_row)
    if ls.cell(dum1,col+1).value is None:
        col=col+3
    else:
        col=col+2
    col1=col-1
    row=row1    
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value

    if tmp is None:
        col=col+1
        tmp= ls.cell(row,col).value
        tmp1= ls.cell(row,col1).value
    
    rpos=4
    rpos1=rpos+12    
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1
    
    row=row1+1
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=5
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1
    
    row=row1+2
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=6
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1

    row=row1+3
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=7
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1

    row=row1+4
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=8
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1
    
    row=row1+5
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=9
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1    
    
    row=row1+6
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=10
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1 
    
    row=row1+7
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=11
    rpos1=rpos+12
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1   
    
    row=row1+8
    tmp= ls.cell(row,col).value
    tmp1= ls.cell(row,col1).value
    rpos=12
    rpos1=rpos+12    
    fs.cell(rpos,cpos).value = tmp
    fs.cell(rpos1,cpos).value = tmp1 
    
   
    print(f'#############  index = {index}')   
    index=index+1
    
fb.save(mainlink+"\\Final Aegon.xlsx")


    
    
    