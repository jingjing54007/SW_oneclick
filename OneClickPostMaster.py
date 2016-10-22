import Tkinter as tk
import tkFont,ttk
import sys,shutil
import telnetlib
import time,re,threading
import platform
import ConfigParser
import tkFileDialog
from FileDialog import *
from win32com.client import Dispatch
from threading import *
import win32com
from Queue import Queue
from  LogCompareOneClick  import LogCompareOneClick



root = tk.Tk()
root.geometry('600x540+100+50')
root.configure()
root.title('One Click Post Master')
root.resizable(False,False)
root.iconbitmap('D:\\Tests\\PyApp\\icon3.ico')




file_t = ''
file_o = ''
var_t = StringVar()
var_o = StringVar()
var_o.set('')
text_line = 1.0


test_set_id = 0
test_set_id_select = "Not Selected"

qc_test_set_d = {}
thread_list = []
msg_q = Queue()
#-------------------------------------------------------------------------
def monitor(root            
            ):
    for each in thread_list:
        if not each[0].isAlive():
            print each
            each[1]['state']='normal'
            thread_list.remove(each)

##    if not msg_q.empty():
##        displayInfo(msg_q.get())
   

    root.after(1000,
               monitor,
               root
               
               )
def displayInfo(msg):
    global text_line
    text_line = text_line + 1.0
    text.configure(state='normal')
    text.insert(text_line, msg)
    text.configure(state='disabled')
    text.see(tk.END)

    
    
def openFile(file_type):
    file_m = tkFileDialog.askopenfilename(initialdir = os.path.dirname(r"D:\\"),
                                          filetypes=[('Excel file',('*.xls','*xlsx'))]
                                          )
    if file_type == "temp":
        var_t.set(file_m)
        displayInfo("Temp result file selected: %s\n" % file_m)
    elif file_type == "official":
        var_o.set(file_m)
        displayInfo("Official report file selected: %s\n" % file_m)
#---------------------------------------------------------------------------------
def test():
    for i in range(0,10):
        displayInfo("%d" %i)
        time.sleep(1000)
def GO(task,buttn):
    global thread_list
    if task == "Merge report":
        myThread = Thread(target = mergeReport,args=[buttn,])
        myThread.start()
    elif task == "Copy log":
        path_log = ""
        path_log = tkFileDialog.askdirectory(initialdir=r"Y:\R&D_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\Intel",title='Please select a directory')
        if len(path_log ) == 0:
            displayInfo('Wrong log path\n')
            return        
        myThread = Thread(target = copyLog,args=[path_log,])
        myThread.start()
        buttn['state']='disable'
        thread_list.append((myThread,buttn))
    elif task == "Compare log":
        path_log_this = tkFileDialog.askdirectory(initialdir=r"Y:\R&D_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\Intel",
                                                  title='Select log path for this run')
        path_log_last = tkFileDialog.askdirectory(initialdir=path_log_this,#r"Y:\R&D_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\Intel",
                                                  title='Select log path for last run')
        myThread = Thread(target = compareLog,args=[path_log_this,path_log_last,])
        myThread.start()        
        buttn['state']='disable'
        thread_list.append((myThread,buttn))


        
    
###########################################################################
def findCol(sh,name):
    i = 0
    for col in range(1,sh.UsedRange.Columns.Count+1):
        if sh.Cells(1,col).Value == name:
            i = col
    return i

def findCol2(sh,name):
    i = 0
    for col in range(1,sh.UsedRange.Columns.Count+1):
        if sh.Cells(2,col).Value == name:
            i = col
    return i
#--------------------------------------------------------------------------
def getTestSet(filewin,tree,buttn):    
    import pythoncom
    pythoncom.CoInitialize()
    global test_set_id
    global test_set_id_select
    item = tree.selection()
    values = tree.item(item,"values")
    displayInfo("Test Set Name: %s\n" %values[1])
    displayInfo("Test Set ID: %s\n" %values[2])

    test_set_id_select = values[2]
    
    filewin.destroy()
    myThread = Thread(target = importQC,args=[values[1],values[2],])
    myThread.start()
    #buttn['state']='disable'
    thread_list.append((myThread,buttn))
        
def importQC(set_name, set_id):
    try:
        import pythoncom
        pythoncom.CoInitialize()
        
        if not os.path.isfile( var_o.get() ):
            displayInfo("1-Click Official Report is not selected !!!\n")
            return

        qcServer = "http://frilm-hpalm:8080/qcbin/"
        qcUser = "oneclick"
        qcPassword = "sierra_211"
        qcDomain = "DEFAULT"
        qcProject = "Validation"

        from datetime import datetime

        displayInfo( "QC login ...\n")
        
        t = win32com.client.Dispatch("TDApiOle80.TDConnection")
        t.InitConnectionEx(qcServer)
        t.Login(qcUser,qcPassword)
        t.Connect(qcDomain,qcProject)

        displayInfo( "QC Logged in\n")
        

        mg=t.TreeManager
        npath=r"Root\AT\INTEL"
        #npath=r"Root\Temp\RTN"
        tsFolder = t.TestSetTreeManager.NodeByPath(npath)

        tfactory=tsFolder.TestSetFactory
        td_tsff=tfactory.Filter
        td_testset=td_tsff.NewList()

        displayInfo("Looking for Test Campaign : %s, %s\n" % (set_name,set_id))

        TSetFact = t.TestSetFactory
        tsFilter = TSetFact.Filter
        tsFilter.SetFilter("CY_CYCLE_ID","%s" % set_id) #7618: AT_HL85xx_Centaurus
        tsList = tsFolder.FindTestSets("",False,tsFilter.Text)    
        
        otest = tsList.Item(1)
        td_TSTestSetFactory = otest.TSTestFactory
        #add filter
        tetsFilter = td_TSTestSetFactory.Filter
        tetsFilter.SetFilter('TC_ACTUAL_TESTER',"oneclick")#tester == rtn
        tetsFilter.SetFilter('TC_STATUS','"No Run" Or "Replay"')#Status == Replay
        td_tstsff = td_TSTestSetFactory.NewList(tetsFilter.Text)


        displayInfo( "Export TCs list...\n")

        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True

        wb= xlApp.Workbooks.Open(Filename= var_o.get())
        sh = wb.sheets["Sheet1"]

        d = {}
        tcs = {}
        col_TestName = findCol2(sh,'Test Name')
        col_ModuleType = findCol2(sh,'ModuleType')
        col_ModuleRef =  findCol2(sh,'ModuleRef')      
        col_SIM = findCol2(sh,'SIM')
        
        col_Result = 1    
        for col in range(1,sh.UsedRange.Columns.Count+1):
            if str(sh.Cells(1,col).Value) != 'None':
                col_Result = col

        displayInfo("Result in column %d will be impoted\n" % col_Result)
        
        col_Issue = col_Result + 1
        col_QC_Posted = col_Result + 2
        
        
        for row in range(3,sh.UsedRange.Rows.Count+1):
            if str(sh.Cells(row, col_QC_Posted).Value) == "Need Import to QC":                
                tcs[sh.Cells(row, col_TestName).Value] = {'Result' : sh.Cells(row, col_Result).Value,
                                                          'ModuleRef' : sh.Cells(row, col_ModuleRef).Value,
                                                          'ModyleType' : sh.Cells(row, col_ModuleType).Value,
                                                          'SIM' : sh.Cells(row, col_SIM).Value,
                                                          'IssueID' : str(sh.Cells(row, col_Issue).Value).replace("-Active",""),
                                                          'Row':row
                                                          }
        VersionSoft = sh.Cells(1, col_Result).Value
              
        for otestitem in td_tstsff:        
            if otestitem.TestName in tcs.keys():
                if '[1]' not in otestitem.Name:
                    continue
                displayInfo( "%s %s Importing ...\n" % (str(datetime.now().strftime("%y/%m/%d %H:%M:%S")), str(otestitem.Name)))
                #msg_q.put("\n\n%s %s Importing ...\n " % (str(datetime.now().strftime("%y/%m/%d %H:%M:%S")), str(otestitem.Name)))

                td_RunFactory = otestitem.RunFactory
                obj_theRun = td_RunFactory.AddItem("Run_" + datetime.now().strftime("%m-%d_%H-%M-%S"))
                obj_theRun.Status = tcs[otestitem.TestName]['Result']
                obj_theRun.SetField('RN_USER_01',VersionSoft) #VersionSoft
                obj_theRun.SetField('RN_USER_17',tcs[otestitem.TestName]['IssueID'].replace("None",'')) #Issue ID
                obj_theRun.SetField('RN_USER_02',"N/A") #Flash Tyep
                obj_theRun.SetField('RN_USER_03',"%s" % tcs[otestitem.TestName]['ModyleType']) #Module Tyep
                obj_theRun.SetField('RN_USER_04',"%s" % tcs[otestitem.TestName]['ModuleRef']) #Module Ref
                obj_theRun.SetField('RN_USER_08',"%s" % tcs[otestitem.TestName]['SIM']) #SIM
                obj_theRun.SetField('RN_USER_09',"1-click") #Test Equipment
                obj_theRun.SetField('RN_USER_10',"Not Applicable") #Framework version
                obj_theRun.SetField('RN_TESTER_NAME',"oneclick")#Tester                
                
                if '_AVMS_' in otestitem.TestName:
                    obj_theRun.SetField('RN_USER_12',"AvMS 4.x") #WDMS version              

                if otestitem.IsLocked:
                    otestitem.UnLockObject()

                #otestitem.SetField('TC_USER_01',"") #Comment
                otestitem.SetField('TC_USER_19',tcs[otestitem.TestName]['IssueID'].replace("None",'')) #Issue ID
                otestitem.SetField('TC_USER_02',VersionSoft) #VersionSoft
                otestitem.SetField('TC_USER_03',tcs[otestitem.TestName]['ModyleType']) #Module Type
                otestitem.SetField('TC_USER_04',"N/A") #Flash Type
                otestitem.SetField('TC_USER_05',tcs[otestitem.TestName]['ModuleRef']) #Module Ref
                otestitem.SetField('TC_USER_06',tcs[otestitem.TestName]['SIM']) #sim
                otestitem.SetField('TC_USER_08',"1-click") #Test Equipment
                otestitem.SetField('TC_USER_09',"Not Applicable") #Framework version
                otestitem.SetField('TC_USER_11',"Not Applicable") #Plugin
                otestitem.SetField('TC_ACTUAL_TESTER',"oneclick")#Tester
                if '_AVMS_' in otestitem.TestName:
                    otestitem.SetField('TC_USER_13',"AvMS 4.x") #AVMS server
                else:
                    otestitem.SetField('TC_USER_13',"Not Applicable") #AVMS server
                otestitem.Status = tcs[otestitem.TestName]['Result']
                
                obj_theRun.Post()
                otestitem.Post()

                displayInfo( "%s %s Imported\n" % (str(datetime.now().strftime("%y/%m/%d %H:%M:%S")), str(otestitem.Name)))

                sh.Cells(tcs[otestitem.TestName]['Row'], col_QC_Posted).Value = "Imported"
                time.sleep(1)

        displayInfo( "Imported\n")
        
        t.Logout()

        displayInfo( "QC Logged out\n")
        #buttn['state']='normal'
    except Exception,e:
        displayInfo(type(e))
        displayInfo(e)
        displayInfo("\n--------->Problem on QC import !!!\n\n")
        #buttn['state']='normal'
         
#--------------------------------------------------------------------------------------
def exportQC():
    import pythoncom
    pythoncom.CoInitialize()
    global qc_test_set_d
    qcServer = "http://frilm-hpalm:8080/qcbin/"
    qcUser = "oneclick"
    qcPassword = "sierra_211"
    qcDomain = "DEFAULT"
    qcProject = "Validation"

    from datetime import datetime

    #displayInfo( "QC login ...\n")
    t = win32com.client.Dispatch("TDApiOle80.TDConnection")
    t.InitConnectionEx(qcServer)
    t.Login(qcUser,qcPassword)
    t.Connect(qcDomain,qcProject)

    #displayInfo( "QC Logged in\n")

    mg=t.TreeManager
    npath=r"Root\AT\INTEL"
    #npath=r"Root\Temp\RTN"
    tsFolder = t.TestSetTreeManager.NodeByPath(npath)

    tfactory=tsFolder.TestSetFactory
    td_tsff=tfactory.Filter
    td_testset=td_tsff.NewList()

    TSetFact = t.TestSetFactory
    tsFilter = TSetFact.Filter
    #tsFilter.SetFilter("CY_CYCLE_ID","14321") #7618: AT_HL85xx_Centaurus
    tsList = tsFolder.FindTestSets("",False)

    qc_test_set_d = {}

    for ts in tsList:
        qc_test_set_d[ts.ID] = {}
        qc_test_set_d[ts.ID]['Test Set Name'] = ts.Name
        qc_test_set_d[ts.ID]['Father Folder'] = ts.TestSetFolder
#----------------------------------------------------------------------------------------------
def getQcDirectory(msg, buttn):
    try:
        global qc_test_set_d
        global test_set_id_select
        if len(qc_test_set_d.keys()) == 0:
            displayInfo("QC data base not ready, please wait... and try late\n")
            return
        
        buttn['state']='disable'
        
        filewin = Toplevel()    
        filewin.geometry('595x420+300+10')
        filewin.title('Select Test Campaign to Import')

        tree = ttk.Treeview(filewin,columns=("Nick","Mensaje","Hora"), selectmode="extended")
        tree.heading('#1',text='QC Path',anchor=tk.W)
        tree.heading('#2',text='TestSet Name',anchor=tk.W)
        tree.heading('#3',text='TestSet ID',anchor=tk.W)
        tree.column('#1', stretch=NO, minwidth=0, width=300)
        tree.column('#2', stretch=NO, minwidth=0, width=200)
        tree.column('#3', stretch=NO, minwidth=0, width=65)
        tree.column('#0', stretch=NO, minwidth=0, width=0) #width 0 to not display it
        tree.bind("<Double-1>", lambda x: getTestSet(filewin,tree,buttn))
        ysb = ttk.Scrollbar(filewin,orient='vertical', command=tree.yview)
        ysb.pack()
        ysb.place(height=400,width=20,x=565,y=10)
        tree.configure(yscroll=ysb.set)
        tree.pack()
        tree.place(height=400,width=565,x=10,y=10)
        #for key in qc_test_set_d.keys():
        tag = ['oddrow','evenrow']
        lin = 1
        for key in sorted(qc_test_set_d,key=lambda x:int(x)):
            #tree.insert("", "end", text=r"Root\AT\INTEL\..\%s %s %d" % (d[key]['Father Folder'],d[key]['Test Set Name'], key))
            tree.insert("", "end", values=((r"Root\AT\INTEL\..\%s" % qc_test_set_d[key]['Father Folder'],qc_test_set_d[key]['Test Set Name'], key)),tags =tag[lin%2])
            lin += 1
        tree.tag_configure('oddrow', background='grey')

    except Exception,e:
        displayInfo(type(e))
        displayInfo(e)
        displayInfo("\n--------->Problem on QC export !!!\n\n")
        buttn['state']='normal'



#--------------------------------------------------------------------------
    
def copyLog(path_log):    
    try:
        import pythoncom
        pythoncom.CoInitialize()

        if not os.path.isfile( var_o.get() ):
            displayInfo("1-Click Official Report is not selected !!!\n")
            return
    
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb= xlApp.Workbooks.Open(Filename=var_o.get())    
        sh = wb.sheets["Sheet1"]

        
        if "OneClick" not in path_log:
            ydriver = r"%s\OneClick" % path_log
        else:
            ydriver = path_log
        if not os.path.exists(ydriver):
            os.makedirs(ydriver)

        ydriver = "%s\%%s" % ydriver
        

        logPath = r'Y:\R&D_Product_Enhancement\Common\Validation_APAC\Tests_HK\LOG\1click\%s\Build%s\%s\%s.log'
        col_Result = 0
        print sh.UsedRange.Columns.Count
        for col in range(1,sh.UsedRange.Columns.Count+1):
            #print "----"
            try:
                if str(sh.Cells(1,col).Value) != 'None':
                    col_Result = col
            except:
                print col
        by_TestNumber = False

        displayInfo("Results in %d column\n" % col_Result)
        

        if col_Result == 0:
            return
        
        for row in range(3,sh.UsedRange.Rows.Count+1):
            try:
                if sh.Cells(row, col_Result).Value is None:
                    continue
                if sh.Cells(row, col_Result).Value == 'N/A':
                    continue
                
                link = sh.Range(sh.Cells(row, col_Result).Address).Hyperlinks.Item(1).Address
                #source = logPath % link.split('=.')[1].split('.')[0].replace('/','\\')
                temp = link.split("/")
                source = logPath % (temp[4],temp[5],temp[9],temp[10].split(".")[0])

                if by_TestNumber:              

                    dest = ydriver % (sh.Range(sh.Cells(row, col_Result).Address).Hyperlinks.Item(1).SubAddress.upper() + '.log')           
                    shutil.copyfile(source, dest )
                else:
                            
                    dest = ydriver % (source.split("\\")[-1].upper())
                    
                    temp = dest.split(".LOG")[0] + '.log'
                    dest = temp         

                    if os.path.exists(dest):
                        continue
                    shutil.copyfile(source, dest )
                    displayInfo("%s copied\n" % source)
            
            except Exception, e:
                displayInfo(type(e))
                displayInfo(e)
                displayInfo("\n--------->Problem on log copy !!!\n\n")
                buttn['state']='normal'
        #buttn['state']='normal'
        displayInfo("\n------------End of Log Copy--------------\n")
    except Exception,e:
        displayInfo(type(e))
        displayInfo(e)
        displayInfo("\n--------->Problem on log copy !!!\n\n")
        #buttn['state']='normal'    

def mergeReport(buttn):
    try:
        import pythoncom
        pythoncom.CoInitialize()
        buttn['state']='disable'
        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb = xlApp.Workbooks.Open(Filename=var_t.get())    
        sh = wb.sheets["Sheet1"]

        col_Loop = findCol(sh,"Loop 1")        
        col_IssueID = findCol(sh,"Issue ID")
        col_ModuleType = findCol(sh,"Module Type")
        col_ModuleRef = findCol(sh,"Module Ref")
        col_SIM = findCol(sh,"SIM")
        col_FWver = findCol(sh,"FW Version")
        col_Platform = findCol(sh,"Platform")

        temp_fw_ver = 'hello'
        d = {}
        for row in range(2,sh.UsedRange.Rows.Count+1):
            displayInfo( "Row %d: %s\n" % (row,sh.Cells(row,1).Value))
            d[sh.Cells(row,1).Value] = {'Status':sh.Cells(row,col_Loop).Value,
                                        'Comment': sh.Cells(row,col_IssueID).Value,
                                        'Hyperlink':sh.Range(sh.Cells(row,col_Loop).Address).Hyperlinks.Item(1).Address + '#' + sh.Cells(row,1).Value,
                                        'Module Type':sh.Cells(row,col_ModuleType).Value,
                                        'Module Ref':sh.Cells(row,col_ModuleRef).Value,
                                        'SIM':sh.Cells(row,col_SIM).Value,
                                        'FW Version':str(sh.Cells(row,col_FWver).Value),
                                        'PlatForm':sh.Cells(row,col_Platform).Value
                                        }
            if temp_fw_ver == 'hello':
                if str(sh.Cells(row,col_FWver).Value) not in ['None','Unkown Version'] and '\r' not in str(sh.Cells(row,col_FWver).Value) and '\n' not in str(sh.Cells(row,col_FWver).Value):
                    temp_fw_ver = str(sh.Cells(row,col_FWver).Value)

        wb.Close()

        displayInfo( '\n------------------------------------------------------------------\n'  )      

        wb2= xlApp.Workbooks.Open(Filename=var_o.get())
        sh2 = wb2.sheets["Sheet1"]

        row_end = sh2.UsedRange.Rows.Count+1
        if row_end == 4:
            row_end-=1

        col_Status = 0
        for col in range(3,sh2.UsedRange.Columns.Count+1):
            if sh2.Cells(1,col).Value == temp_fw_ver:
                col_Status = col
                break
        if col_Status == 0:
            col_Status = sh2.UsedRange.Columns.Count+1
        col_IssueID = col_Status+1
        col_Import = col_IssueID+1

        col_ModuleType = findCol2(sh2,"ModuleType")
        col_ModuleRef = findCol2(sh2,"ModuleRef")
        col_SIM = findCol2(sh2,"SIM")
        col_PlatForm = findCol2(sh2,"1-Click")

        sh2.Cells(2,col_Status).Value = "Result"
        sh2.Cells(2,col_IssueID).Value = "Issue ID"
        sh2.Cells(2,col_Import).Value = "Import QC"

        sh2.Cells(1,col_Status).Value = temp_fw_ver
        
        for row in range(3,row_end):
            if sh2.Cells(row,1).Value in d.keys():
                sh2.Cells(row,col_Status).Value = d[sh2.Cells(row,1).Value]['Status']
                sh2.Cells(row,col_IssueID).Value = d[sh2.Cells(row,1).Value]['Comment']
                sh2.Hyperlinks.Add(Anchor = sh2.Range(sh2.Cells(row,col_Status).Address),Address = d[sh2.Cells(row,1).Value]["Hyperlink"])
                sh2.Cells(row,col_ModuleType).Value = d[sh2.Cells(row,1).Value]['Module Type']
                sh2.Cells(row,col_ModuleRef).Value = d[sh2.Cells(row,1).Value]['Module Ref']
                sh2.Cells(row,col_SIM).Value = d[sh2.Cells(row,1).Value]['SIM']
                sh2.Cells(row,col_PlatForm).Value = d[sh2.Cells(row,1).Value]['PlatForm']
                sh2.Cells(row,col_Import).Value = (lambda x: "Need Import to QC" if x == "Passed" else None)(str(d[sh2.Cells(row,1).Value]['Status']))
##                if d[sh2.Cells(row,1).Value]['FW Version'] not in ['None','Unknown','Unkown Version']:
##                    sh2.Cells(1,col_Status).Value = d[sh2.Cells(row,1).Value]['FW Version']
                del d[sh2.Cells(row,1).Value]
                displayInfo("Row %d: %s @ Temp report\n" % (row,sh2.Cells(row,1).Value))

        for key in sorted(d.keys()):
            displayInfo("Row %d: %s @ Official report\n" % (row_end,key))        
            sh2.Cells(row_end,1).Value = key
            sh2.Cells(row_end,col_Status).Value     = d[sh2.Cells(row_end,1).Value]['Status']
            sh2.Cells(row_end,col_IssueID).Value    = d[sh2.Cells(row_end,1).Value]['Comment']
            sh2.Hyperlinks.Add(Anchor = sh2.Range(sh2.Cells(row_end,col_Status).Address),Address = d[sh2.Cells(row_end,1).Value]["Hyperlink"])            
            sh2.Cells(row_end,col_ModuleType).Value = d[sh2.Cells(row_end,1).Value]['Module Type']
            sh2.Cells(row_end,col_ModuleRef).Value  = d[sh2.Cells(row_end,1).Value]['Module Ref']
            sh2.Cells(row_end,col_SIM).Value        = d[sh2.Cells(row_end,1).Value]['SIM']
            sh2.Cells(row_end,col_PlatForm).Value        = d[sh2.Cells(row_end,1).Value]['PlatForm']
            sh2.Cells(row_end,col_Import).Value = (lambda x: "Need Import to QC" if x == "Passed" else None)(str(d[sh2.Cells(row_end,1).Value]['Status']))
##            if d[sh2.Cells(row_end,1).Value]['FW Version'] not in ['None','Unknown','Unkown Version']:
##                    sh2.Cells(1,col_Status).Value = d[sh2.Cells(row_end,1).Value]['FW Version']
            row_end += 1

        displayInfo("\n------------End of Report merger--------------\n")
        buttn['state']='normal'
    except Exception,e:
        displayInfo(type(e))
        displayInfo(e)
        displayInfo("\n--------->Problem on report merge !!!\n\n")
        buttn['state']='normal'
#--------------------------------------------------------------------------------
def compareLog(path_log_this,path_log_last):
    try:
        
        logcmp = LogCompareOneClick(path_log_this.replace('/','\\'), path_log_last.replace('/','\\'))
        logcmp.runCompare()
        logcmp.processOneClickResultsByTest()
        d_ref = logcmp.getOneClickList()       

        if not os.path.isfile( var_o.get() ):
            displayInfo("1-Click Official Report is not selected !!!\n")
            return

        xlApp = Dispatch("Excel.Application")
        xlApp.Visible = True
        wb= xlApp.Workbooks.Open(Filename=var_o.get())
        sh = wb.sheets["Sheet1"]

        col_ToAssign = sh.UsedRange.Columns.Count+1
        col_Detail = col_ToAssign + 1
        sh.Cells(2,col_ToAssign).Value = "ToAssign"
        sh.Cells(2,col_Detail).Value = "Detail"

        col_Import = col_ToAssign - 1
        col_Issue = col_Import - 1
        
        for row in range(3,sh.UsedRange.Rows.Count+1):
            displayInfo( "Row %d\n" % row)
            temp = [v for k,v in d_ref.iteritems() if sh.Cells(row,1).Value in k]
            displayInfo( temp)
            displayInfo( "\n")
            sh.Cells(row,col_ToAssign).Value = (lambda x : len(x)==1 and x[0][0] or None)(temp)
            sh.Cells(row,col_Detail).Value = (lambda x : len(x)==1 and x[0][1] or None)(temp)

            if str(sh.Cells(row,col_ToAssign).Value) == "No" and "-Active" in str(sh.Cells(row,col_Issue).Value):
                sh.Cells(row,col_Import).Value = "Need Import to QC"
            time.sleep(1)

    except Exception,e:
        displayInfo(type(e))
        displayInfo(e)
        displayInfo("\n--------->Problem on log compare !!!\n\n")        

        
#--------------------------------------------------------------------------------

from PIL import ImageTk, Image
im = Image.open('D:\\Tests\\PyApp\\iot3.png')
tkimage = ImageTk.PhotoImage(im)
label_i = tk.Label(root,image=tkimage)
label_i.pack()
label_i.place(height=160,width=600,x=0,y=0)

##frame_1 = tk.LabelFrame(label_i,
##                        text = "Make Official Report"
##                        )
##frame_1.pack(fill="both", expand="yes")
##frame_1.place(height=130,width=580,x=10,y=10)



button_t = tk.Button(label_i, text ="Select Temp Report", command = lambda :openFile("temp"))
button_t.pack()
button_t.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"))
button_t.place(height=40,width=180,x=15,y=40)

button_o = tk.Button(label_i, text ="Select Official Report", command = lambda :openFile("official"))
button_o.pack()
button_o.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"))
button_o.place(height=40,width=180,x=205,y=40)

button_m = tk.Button(label_i, text ="Merge Report", command = lambda :GO("Merge report",button_m))
button_m.pack()
button_m.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_m.place(height=40,width=180,x=395,y=40)

button_c = tk.Button(label_i, text ="Copy Log Files", command = lambda :GO("Copy log",button_c))
button_c.pack()
button_c.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_c.place(height=40,width=180,x=15,y=90)

button_i = tk.Button(label_i, text ="Import QC", command = lambda :getQcDirectory("Import QC",button_i))
button_i.pack()
button_i.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_i.place(height=40,width=180,x=205,y=90)

button_cp = tk.Button(label_i, text ="Compare Log", command = lambda :GO("Compare log",button_cp))
button_cp.pack()
button_cp.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_cp.place(height=40,width=180,x=395,y=90)



##button_1.pack()
##button_1.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
##button_1.place(height=40,width=60,x=80,y=10)

probar_1 = ttk.Progressbar(label_i,
                           orient ="horizontal",
                           mode ="determinate"
                           )


#--------------------------------------------------------------------------------
scrollbar = tk.Scrollbar(root)
scrollbar.pack()
scrollbar.place(height=360,width=20,x=570,y=165)

text = tk.Text(root)
text.pack()
text.configure(background='black',foreground='white')
text.configure(state='disabled')
text.place(bordermode=OUTSIDE,height=360,width=560,x=10,y=165)
text.config(yscrollcommand=scrollbar.set)
scrollbar.config(command=text.yview)


myThread = Thread(target = exportQC)
myThread.start()
####################################################################################
root.after(1000,
           monitor,
           root
           
           )


root.mainloop()
