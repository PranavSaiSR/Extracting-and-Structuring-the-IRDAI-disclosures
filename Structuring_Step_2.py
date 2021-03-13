# -*- coding: utf-8 -*-
"""
Created on Thu Jul 16 15:56:41 2020

@author: PRANAVSAI
"""

# -*- coding: utf-8 -*-
"""
Created on Sun May 31 13:24:20 2020

@author: PRANAVSAI
"""


################################
################################
################################
################################
################################
fb = load_workbook(mainlink+"\\Final Shriram.xlsx")

fs = fb.get_sheet_by_name("Sheet1")




index=1
a=[]
for ind in range(len(excel)):
    exc = excel[index]
    
    lb = load_workbook(mainlink+"/"+exc)
    
    
    ls= lb.get_sheet_by_name(str(1))  
    
    if index<200:
        dum1,col=wordfinder("L-4",1,ls.max_row)
        col=col+1
        if ls.cell(dum1,col).value is None:
            col=col+1
            if ls.cell(dum1,col).value is None:
                col=col+1      
    else:
        dum1,col=wordfinder("Grand Total",1,ls.max_row)
        col=col
    
    row1,dum=wordfinder("Premiums earned - net",1,ls.max_row)   
    row=row1+1
    if ls.cell(row,col).value is None:
        col=col+1
    
    tmp= ls.cell(row,col).value
    a.append(tmp)   
    rpos=2
    cpos=colind(exc,2)
    fs.cell(rpos,cpos).value = tmp

    
    
    row=row1+2
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=3
    cpos=colind(exc,1)    
    fs.cell(rpos,cpos).value = tmp
        
    row=row1+3
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=4
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    # row,dum=wordfinder("Operating expenses related to insurance business",1,ls.max_row)
    row,dum=wordfinder("Commission",1,ls.max_row)
    row=row
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=6
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row=row+1
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=7
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row,dum=wordfinder("Benefits Paid (Net)",1,ls.max_row) 
    row=row
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=8
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row,dum=wordfinder("SURPLUS/ (DEFICIT)  (D) =(A)-(B)-(C)",row,ls.max_row) 
    row=row
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=9
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp




    ls= lb.get_sheet_by_name(str(2))       
    row,dum=wordfinder("Provision for Taxation  ",1,ls.max_row) 
    row=row-1
    if ls.cell(row,dum+1).value is None:        
        col=dum+2
    else:
        col=dum+1
    if ls.cell(row,dum+2).value is None and ls.cell(row,dum+1).value is None:
        col= dum+3
    tmp=ls.cell(row,col).value
    a.append(tmp)
    rpos=10
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp    
        
    
    
    ls= lb.get_sheet_by_name(str(3))
    dum,col=wordfinder("Share Capital",1,ls.max_row)
    # col=col+1
    
    if ls.cell(dum,col+1).value is None:
        col=col+3
        if ls.cell(dum,col).value is None:
            col=col+1
    else:
        col=col+2
        if ls.cell(dum,col).value is None:
            col=col+1    
    
    row,dum=wordfinder("Share Capital",1,ls.max_row) 
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=12
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp 
           
    row,dum=wordfinder("Reserves And Surplus",1,ls.max_row) 
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=13
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp     
    
    row,dum=wordfinder("Borrowings",1,ls.max_row) 
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=14
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row1,dum=wordfinder("APPLICATION OF FUNDS",1,ls.max_row)
    row1=row1
    row=row1+2
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=15
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row=row1+3
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=16
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
       
    row=row1+4
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=17
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp
    
    row=row1+5
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=18
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp   
    
    row=row1+6
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=19
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp 
    
    row1,dum=wordfinder("Cash and Bank Balances",1,ls.max_row)
    
    row=row1
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=20
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp  
    
    row=row1+1
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=21
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp   
    
    
    row=row1+3
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=23
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp  
    
    row=row1+4
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=24
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp   
    
    row=row1+6
    tmp= ls.cell(row,col).value
    if tmp is None:
        tmp= ls.cell(row,col+1).value
    a.append(tmp)
    rpos=26
    cpos=colind(exc,1)
    fs.cell(rpos,cpos).value = tmp 
    
    
    
    # ls= lb.get_sheet_by_name(str(5))   
    # row1,dum=wordfinder("Solvency Ratio (ASM/RSM)",1,ls.max_row)
    
    # col=dum+1
    # if ls.cell(row1,col).value is None:
    #     col=col+1
    #     if ls.cell(row1,col).value is None:
    #         col=col+1
    #         if ls.cell(row1,col).value is None:
    #             col=col+1
    #             if ls.cell(row1,col).value is None:
    #                 col=col+1
    # row=row1-2
    # tmp= ls.cell(row,col).value
    # if tmp is None:
    #     tmp= ls.cell(row,col+1).value
    # a.append(tmp)
    # rpos=28
    # cpos=colind(exc,1)
    # fs.cell(rpos,cpos).value = tmp     
    
    # row=row1-1
    # tmp= ls.cell(row,col).value
    # if tmp is None:
    #     tmp= ls.cell(row,col+1).value
    # a.append(tmp)
    # rpos=29
    # cpos=colind(exc,1)
    # fs.cell(rpos,cpos).value = tmp   

    # row=row1
    # tmp= ls.cell(row,col).value
    # if tmp is None:
    #     tmp= ls.cell(row,col+1).value
    # a.append(tmp)
    # rpos=30
    # cpos=colind(exc,1)
    # fs.cell(rpos,cpos).value = tmp
    
    
    

    print(f'#############  index = {index}')
    index=index+1





fb.save(mainlink+"\\Final Shriram.xlsx")
   
    

          