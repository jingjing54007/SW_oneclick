import Tkinter as tk
import tkFont,ttk
import sys
import telnetlib
import time,re,threading
import platform
import ConfigParser
from ReliabilityTest import *

root = tk.Tk()
root.geometry('880x540+100+50')
root.configure()
root.title('1-Click Team Stress Test System @V2.1')
root.resizable(False,False)
##L1 = tk.Label(root, text="APC IP Addr")
##L1.pack()
##L1.place(height=30,width=70,x=10,y=10)

#-------------------------------------------------------------------------
def monitor(root,
            probar_1,
            probar_2,
            probar_3,
            probar_4,
            probar_5,
            probar_6,
            probar_7,
            probar_8,
            probar_9,
            probar_10
            ):
    probar_1['maximum'] = script_number[1]
    probar_1['value'] = progress[1]
    probar_2['maximum'] = script_number[2]
    probar_2['value'] = progress[2]
    probar_3['maximum'] = script_number[3]
    probar_3['value'] = progress[3]
    probar_4['maximum'] = script_number[4]
    probar_4['value'] = progress[4]
    probar_5['maximum'] = script_number[5]
    probar_5['value'] = progress[5]
    probar_6['maximum'] = script_number[6]
    probar_6['value'] = progress[6]
    probar_7['maximum'] = script_number[7]
    probar_7['value'] = progress[7]
    probar_8['maximum'] = script_number[8]
    probar_8['value'] = progress[8]
    probar_9['maximum'] = script_number[9]
    probar_9['value'] = progress[9]
    probar_10['maximum'] = script_number[10]
    probar_10['value'] = progress[10]

    root.after(1000,
               monitor,
               root,
               probar_1,
               probar_2,
               probar_3,
               probar_4,
               probar_5,
               probar_6,
               probar_7,
               probar_8,
               probar_9,
               probar_10
               )

###########################################################################
#--------------------------------------------------------------------------------
frame_1 = tk.LabelFrame(root,
                 text = "DUT #1",)
frame_1.pack(fill="both", expand="yes")
frame_1.place(height=80,width=420,x=10,y=10)

com_value = tk.StringVar()
combox_1 =  ttk.Combobox(frame_1,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_1.pack()
combox_1.place(height=40,width=60,x=10,y=10)

button_1 = tk.Button(frame_1,
                     text = "Start",
                     command = lambda :GO(1,button_1)
                     )
combox_1.current(0)
button_1.pack()
button_1.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_1.place(height=40,width=60,x=80,y=10)

probar_1 = ttk.Progressbar(frame_1,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_1.pack()
probar_1.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_2 = tk.LabelFrame(root,
                 text = "DUT #2")
frame_2.pack(fill="both", expand="yes")
frame_2.place(height=80,width=420,x=10,y=100)

com_value = tk.StringVar()
combox_2 =  ttk.Combobox(frame_2,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_2.current(1)
combox_2.pack()
combox_2.place(height=40,width=60,x=10,y=10)

button_2 = tk.Button(frame_2,
                     text = "Start",
                     command = lambda :GO(2,button_2)
                     )
button_2.pack()
button_2.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_2.place(height=40,width=60,x=80,y=10)

probar_2 = ttk.Progressbar(frame_2,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_2.pack()
probar_2.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_3 = tk.LabelFrame(root,
                 text = "DUT #3")
frame_3.pack(fill="both", expand="yes")
frame_3.place(height=80,width=420,x=10,y=190)

com_value = tk.StringVar()
combox_3 =  ttk.Combobox(frame_3,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_3.current(2)
combox_3.pack()
combox_3.place(height=40,width=60,x=10,y=10)

button_3 = tk.Button(frame_3,
                     text = "Start",
                     command = lambda :GO(3,button_3)
                     )
button_3.pack()
button_3.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_3.place(height=40,width=60,x=80,y=10)

probar_3 = ttk.Progressbar(frame_3,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_3.pack()
probar_3.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_4 = tk.LabelFrame(root,
                 text = "DUT #4")
frame_4.pack(fill="both", expand="yes")
frame_4.place(height=80,width=420,x=10,y=280)

com_value = tk.StringVar()
combox_4 =  ttk.Combobox(frame_4,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_4.current(3)
combox_4.pack()
combox_4.current(1)
combox_4.place(height=40,width=60,x=10,y=10)

button_4 = tk.Button(frame_4,
                     text = "Start",
                     command = lambda :GO(4,button_4)
                     )
button_4.pack()
button_4.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_4.place(height=40,width=60,x=80,y=10)

probar_4 = ttk.Progressbar(frame_4,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_4.pack()
probar_4.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_5 = tk.LabelFrame(root,
                 text = "DUT #5")
frame_5.pack(fill="both", expand="yes")
frame_5.place(height=80,width=420,x=10,y=370)

com_value = tk.StringVar()
combox_5 =  ttk.Combobox(frame_5,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_5.current(4)
combox_5.pack()
combox_5.place(height=40,width=60,x=10,y=10)

button_5 = tk.Button(frame_5,
                     text = "Start",
                     command = lambda :GO(5,button_5)
                     )
button_5.pack()
button_5.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_5.place(height=40,width=60,x=80,y=10)

probar_5 = ttk.Progressbar(frame_5,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_5.pack()
probar_5.place(height=40,width=260,x=150,y=10)

#--------------------------------------------------------------------------------
frame_6 = tk.LabelFrame(root,
                 text = "DUT #6")
frame_6.pack(fill="both", expand="yes")
frame_6.place(height=80,width=420,x=450,y=10)

com_value = tk.StringVar()
combox_6 =  ttk.Combobox(frame_6,
                         #textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3','COM4','COM5','COM6')
                         )
combox_6.current(5)
combox_6.pack()
combox_6.place(height=40,width=60,x=10,y=10)

button_6 = tk.Button(frame_6,
                     text = "Start",
                     command = lambda :GO(6,button_6)
                     )
button_6.pack()
button_6.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_6.place(height=40,width=60,x=80,y=10)

probar_6 = ttk.Progressbar(frame_6,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_6.pack()
probar_6.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_7 = tk.LabelFrame(root,
                 text = "DUT #7")
frame_7.pack(fill="both", expand="yes")
frame_7.place(height=80,width=420,x=450,y=100)

com_value = tk.StringVar()
combox_7 =  ttk.Combobox(frame_7,
                         textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3')
                         )
combox_7.pack()
combox_7.place(height=40,width=60,x=10,y=10)

button_7 = tk.Button(frame_7,
                     text = "Start",
                     command = lambda :GO(7,button_7)
                     )
button_7.pack()
button_7.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_7.place(height=40,width=60,x=80,y=10)

probar_7 = ttk.Progressbar(frame_7,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_7.pack()
probar_7.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_8 = tk.LabelFrame(root,
                 text = "DUT #8")
frame_8.pack(fill="both", expand="yes")
frame_8.place(height=80,width=420,x=450,y=190)

com_value = tk.StringVar()
combox_8 =  ttk.Combobox(frame_8,
                         textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3')
                         )

combox_8.pack()
combox_8.place(height=40,width=60,x=10,y=10)

button_8 = tk.Button(frame_8,
                     text = "Start",
                     command = lambda :GO(8,button_8)
                     )
button_8.pack()
button_8.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_8.place(height=40,width=60,x=80,y=10)

probar_8 = ttk.Progressbar(frame_8,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_8.pack()
probar_8.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_9 = tk.LabelFrame(root,
                 text = "DUT #9")
frame_9.pack(fill="both", expand="yes")
frame_9.place(height=80,width=420,x=450,y=280)

com_value = tk.StringVar()
combox_9 =  ttk.Combobox(frame_9,
                         textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3')
                         )
combox_9.pack()
combox_9.place(height=40,width=60,x=10,y=10)

button_9 = tk.Button(frame_9,
                     text = "Start",
                     command = lambda :GO(9,button_9)
                     )
button_9.pack()
button_9.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_9.place(height=40,width=60,x=80,y=10)

probar_9 = ttk.Progressbar(frame_9,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_9.pack()
probar_9.place(height=40,width=260,x=150,y=10)

#------------------------------------------------------------------------------
frame_10 =tk.LabelFrame(root,
                 text = "DUT #10")
frame_10.pack(fill="both", expand="yes")
frame_10.place(height=80,width=420,x=450,y=370)

com_value = tk.StringVar()
combox_10=  ttk.Combobox(frame_10,
                         textvariable = com_value,
                         state='readonly',
                         values = ('COM1','COM2','COM3')
                         )
combox_10.pack()
combox_10.place(height=40,width=60,x=10,y=10)

button_10 =tk.Button(frame_10,
                     text = "Start",
                     command = lambda :GO(10,button_10)
                     )
button_10.pack()
button_10.configure(font=tkFont.Font(family="Helvetica",size=13,weight="bold"),background = 'magenta',foreground = 'white')
button_10.place(height=40,width=60,x=80,y=10)

probar_10 =ttk.Progressbar(frame_10,
                           orient ="horizontal",
                           mode ="determinate"
                           )

probar_10.pack()
probar_10.place(height=40,width=260,x=150,y=10)
##############################################################################

#------------------------------------------------------------------------------
frame_11 = tk.LabelFrame(root)
                         #text = "Admin")
frame_11.pack(fill="both", expand="yes")
frame_11.configure(background = 'cyan')
frame_11.place(height=65,width=860,x=10,y=460)

button_11 = tk.Button(frame_11,
                      text = "Download Firmware"
                      #command = lambda :GO(5,button_5)
                      )
button_11.pack()
button_11.configure(font=tkFont.Font(family="Helvetica",size=11,weight="bold"))
button_11.place(height=40,width=160,x=10,y=10)

button_11 = tk.Button(frame_11,
                      text = "Create Report"
                      #command = lambda :GO(5,button_5)
                      )
button_11.pack()
button_11.configure(font=tkFont.Font(family="Helvetica",size=11,weight="bold"))
button_11.place(height=40,width=120,x=180,y=10)

button_12 = tk.Button(frame_11,
                      text = "Copy Log Files To"
                      #command = lambda :GO(5,button_5)
                      )
button_12.pack()
button_12.configure(font=tkFont.Font(family="Helvetica",size=11,weight="bold"))
button_12.place(height=40,width=150,x=310,y=10)

####################################################################################
root.after(1000,
           monitor,
           root,
           probar_1,
           probar_2,
           probar_3,
           probar_4,
           probar_5,
           probar_6,
           probar_7,
           probar_8,
           probar_9,
           probar_10
           )


root.mainloop()
