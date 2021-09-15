from tkinter import *
from tkinter import ttk
import fmvreport as fr

root = Tk()
root.title('FMV Report')
root.minsize(680,345)
root.maxsize(680,345)
panel1 = Frame(root, bd=1, relief="raised", bg="#545454")
panel1.pack(fill=BOTH,expand=1)
label1 = Label(panel1, text="R-L Report",bg="#545454",fg="#ffffff",font=("Calibri", 20)).grid(row=0,pady=10)
button1 = Button(panel1,width=20,text='1.Excel(RL1)',bg="#545454",fg="#ffffff", command=fr.excelfilecommand1).grid(row=1,column=0,padx=10,pady=15)
button2 = Button(panel1,width=20,text='2.Excel(RL2)',bg="#545454",fg="#ffffff", command=fr.excelfilecommand2).grid(row=1,column=1,padx=10,pady=15)
button3 = Button(panel1,width=20,text='Çıktı Klasörü',bg="#545454",fg="#ffffff", command=fr.outputfoldercommand).grid(row=1,column=3,padx=10,pady=15)
button4 = Button(panel1,width=20,text='Başlat',bg="#545454",fg="#ffffff", command=fr.rlreportbutton).grid(row=2,column=1,padx=10,pady=15)
button5 = Button(panel1,width=20,text='Durdur',bg="#545454",fg="#ffffff", command=root.destroy).grid(row=2,column=2,padx=10,pady=15)


panel2 = Frame(root,bd=1, relief="raised", bg="#545454")
panel2.pack(fill=BOTH,expand=1)
label2 = Label(panel2, text="R-L-S Report",bg="#545454",fg="#ffffff",font=("Calibri", 20)).grid(row=0,column=0,pady=10)
button6 = Button(panel2,width=20,text='1.Excel(RL1)',bg="#545454",fg="#ffffff", command=fr.excelfilecommand1).grid(row=1,column=0,padx=10,pady=15)
button7 = Button(panel2,width=20,text='2.Excel(RL2)',bg="#545454",fg="#ffffff", command=fr.excelfilecommand2).grid(row=1,column=1,padx=10,pady=15)
button8 = Button(panel2,width=20,text='3.Excel(S1)',bg="#545454",fg="#ffffff", command=fr.excelfilecommand3).grid(row=1,column=2,padx=10,pady=15)
button9 = Button(panel2,width=20,text='Çıktı Klasörü',bg="#545454",fg="#ffffff", command=fr.outputfoldercommand).grid(row=1,column=3,padx=10,pady=15)
button10 = Button(panel2,width=20,text='Başlat',bg="#545454",fg="#ffffff", command=fr.rlsreportbutton).grid(row=2,column=1,padx=10,pady=15)
button11 = Button(panel2,width=20,text='Durdur',bg="#545454",fg="#ffffff", command=root.destroy).grid(row=2,column=2,padx=10,pady=15)


style = ttk.Style()
style.theme_create('Cloud', parent="classic", settings={
    ".": {
        "configure": {
            "background": '#aeb0ce', # All colors except for active tab-button
        }
    },
    "TNotebook": {
        "configure": {
            "background":"#545454", # color behind the notebook
            "tabmargins": [3, 3, 0, 0], # [left margin, upper margin, right margin, margin beetwen tab and frames]
        }
    },
    "TNotebook.Tab": {
        "configure": {
            "background": '#797979', # Color of non selected tab-button
            "padding": [10, 4], # [space beetwen text and horizontal tab-button border, space between text and vertical tab_button border]
        },
        "map": {
            "background": [("selected", '#545454')], # Color of active tab
            "expand": [("selected", [1, 1, 1, 0])], # [expanse of text]
            "foreground": [("selected", "#ffffff"),("!disabled", "#000000")] 
        }
    }
})
style.theme_use('Cloud')


root.config(background="#545454")
root.mainloop()
