import urllib
from bs4 import BeautifulSoup
from win32com.client import Dispatch

url = r"http://cnhkg-ev-hudson:8080/view/All/job/%s/%s/artifact/oneclick/html/loop.html" #Platform, index

##d = {'HL7_dot1CSB':[44,43],
##     'HL7_dot16':[17]     
##     }

d = {'HL7_dot1CSB':[173]  #72,74,76,77,78,79,80      
     }

if __name__ == u'__main__':

    if 1:

        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb = xlApp.Workbooks.Add()
        sh = wb.sheets["Sheet1"]

        for paltform, index_list in d.items():
            if not index_list:
                continue
            for index in index_list:
                print index
                f = urllib.urlopen(url % (paltform, index))
                h = f.read()
                soup = BeautifulSoup(h)
                table = soup.find("table")
                
                headings = [th.get_text() for th in table.find("tr").find_all("th")]
                datasets = []
                
                for row in table.find_all("tr")[1:]:
                    temp = []
                    for td in row.find_all("td"):                    
                        if td.a is not None:            
                            temp.append({'Result' : td.get_text(), 'Link' : td.a.get('href')})
                        else:
                            temp.append( td.get_text() )           
                        
                    dataset = dict(zip(headings, temp))
                    datasets.append(dataset)

                d = {}
                for data in datasets:
                    temp = data['Test: Test Name']
                    data.pop('Test: Test Name')
                    d[temp] = data

                #print d
                

                col_testName = 1
                col_scriptName = 2
                col_owner = 7
                col_QcStatus = 8
                col_QcComment = 20
                col_Platform = 3
                #col_Jenkin = 12

                col_loop1 = 12
                col_build = 14

                sh.Cells(1, col_testName).Value = "Test Name"
                sh.Cells(1, col_scriptName).Value = "Script Name"
                sh.Cells(1, col_owner).Value = "Owner"
                sh.Cells(1, col_QcStatus).Value = "QC Status"
                sh.Cells(1, col_QcComment).Value = "QC Comment"
                sh.Cells(1, col_Platform).Value = "Platform"
                sh.Cells(1, col_loop1).Value = "Merged Result"
    ##                sh.Cells(1, col_loop2).Value = "Loop2 Result"
    ##                sh.Cells(1, col_loop3).Value = "Loop3 Result"
    ##                sh.Cells(1, col_loop4).Value = "Loop4 Result"
    ##                sh.Cells(1, col_loop5).Value = "Loop5 Result"
                #sh.Cells(1, col_Jenkin).Value = "Jenkins Page"

                test_case_list = d.keys()

                for row in range(2,sh.UsedRange.Rows.Count+1):
                    if sh.Cells(row, col_testName).Value in test_case_list:
                        test_case_list.remove(sh.Cells(row, col_testName).Value)
                        

                row_end = sh.UsedRange.Rows.Count
                row_end += 1
                for key in test_case_list:
                    sh.Cells(row_end,col_testName).Value = key
                    row_end += 1
                    

                for row in range(2,sh.UsedRange.Rows.Count+1):                    
                        if sh.Cells(row, col_testName).Value in d.keys():                        
                            for loop_count in range(1,6):
                                loop_count = str(loop_count)
                                if 'loop' + loop_count not in d[sh.Cells(row, col_testName).Value].keys():
                                    continue
                                
                                if sh.Cells(row,col_loop1).Value == 'Passed':
                                    break
                                elif (sh.Cells(row,col_loop1).Value is None) or\
                                     (sh.Cells(row,col_loop1).Value == 'NoTC' and d[sh.Cells(row, col_testName).Value]['loop' + loop_count]['Result'] in ['Failed','Passed']) or\
                                     (sh.Cells(row,col_loop1).Value == 'Failed' and d[sh.Cells(row, col_testName).Value]['loop' + loop_count]['Result'] in ['Passed']):
                                    try:
                                        sh.Cells(row,col_loop1).Value = d[sh.Cells(row, col_testName).Value]['loop' + loop_count]['Result']                                     
                                        sh.Hyperlinks.Add(Anchor = sh.Range(sh.Cells(row,col_loop1).Address),Address = d[sh.Cells(row, col_testName).Value]['loop' + loop_count]['Link'])
                                    except:
                                        sh.Cells(row,col_loop1).Value = 'NoLog'
                                    sh.Cells(row,col_scriptName).Value = d[sh.Cells(row, col_testName).Value]['Test: Script Name']
                                    sh.Cells(row,col_owner).Value = d[sh.Cells(row, col_testName).Value]['Owner']
                                    sh.Cells(row,col_QcStatus).Value = d[sh.Cells(row, col_testName).Value]['Status']
                                    sh.Cells(row,col_QcComment).Value = d[sh.Cells(row, col_testName).Value]['Comment']
                                    sh.Cells(row,col_Platform).Value = paltform
                                    sh.Cells(row,col_build).Value = index
                                
                                #sh.Cells(row,col_Jenkin).Value = "Browse Jenkins"
                                #sh.Hyperlinks.Add(Anchor = sh.Range(sh.Cells(row,col_Jenkin).Address),Address = url % (paltform, index))                        
                            del d[sh.Cells(row, col_testName).Value]
                    

    if 0:
        import shutil, os
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb= xlApp.Workbooks.Open(Filename=r"D:\temp\Book1.xlsx")
        sh = wb.sheets["Sheet1"]
        
        logPath = r'\\cnhkg-ed-hkva17\%s.log'
               
        col_Result = 12

        ydriver = r"Y:\R&D_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\Intel\HL8548\RHL85xx.5.5.11.0.201412191148.x6250_1\Autotest\%s"

        for row in range(4,sh.UsedRange.Rows.Count+1):
            try:
                if sh.Cells(row, col_Result).Value in ['Passed','Failed']:                
                    link = sh.Range(sh.Cells(row, col_Result).Address).Hyperlinks.Item(1).Address
                    source = logPath % link.split('=.')[1].split('.')[0].replace('/','\\')
                    dest = ydriver % (source.split("\\")[-1].upper())
                    
                    temp = dest.split(".LOG")[0] + '.log'
                    dest = temp
                    

                    if os.path.exists(dest):
                        continue
                    shutil.copyfile(source, dest )
            except:
                print row
                
