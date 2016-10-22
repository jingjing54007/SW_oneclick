# -*- coding: utf-8 -*-
import win32com, win32com.client
from win32com.client import Dispatch
 
if __name__ == "__main__":
    if 1:
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb = xlApp.Workbooks.Open(Filename=r"D:\temp\New folder\ChuckBerry25.xlsx")
        sh = wb.sheets["Sheet1"]

        d = {}
        for row in range(2,1110):
            if sh.Cells(row,12).Value != 'NoLog':
                d[sh.Cells(row,1).Value] = {'Status':sh.Cells(row,12).Value,
                                            'Hyperlink':sh.Range(sh.Cells(row,12).Address).Hyperlinks.Item(1).Address +\
                                            '#'+sh.Cells(row,1).Value,
                                            'Comment': sh.Cells(row,5).Value
                                            }
            print row
            #print d[sh.Cells(row,1).Value]["Hyperlink"]

        wb.Close()

        

        wb2= xlApp.Workbooks.Open(Filename=r"D:\one\ChuckBerry_Report.xlsx")
        sh2 = wb2.sheets["Sheet2"]

        #row_end = 447
        col_result = 48
        col_comment = 49
        for row in range(2,1076):
            if sh2.Cells(row,1).Value in d.keys():
                sh2.Cells(row,col_result).Value = d[sh2.Cells(row,1).Value]['Status']
                sh2.Hyperlinks.Add(Anchor = sh2.Range(sh2.Cells(row,col_result).Address),
                                           Address = d[sh2.Cells(row,1).Value]["Hyperlink"])
                sh2.Cells(row,col_comment).Value = d[sh2.Cells(row,1).Value]['Comment']
                del d[sh2.Cells(row,1).Value]



        

                


                  





                

    
