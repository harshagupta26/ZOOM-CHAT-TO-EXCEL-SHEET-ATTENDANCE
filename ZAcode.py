
#autor @harshagupta
#stared 28.09.20

import pandas as pd
import itertools
import os
import tkinter
from tkinter import messagebox as mb


win=tkinter.Tk()

win.geometry("286x290")
win.title("Attendance Generator ")
L1=tkinter.Label(win,text="Zoom Chat Path",fg="black", bg="yellow")
L1.grid(row=0,column=0)
L2=tkinter.Label(win,text="Excel sheet Path",fg="black", bg="yellow")
L2.grid(row=1,column=0)

E1=tkinter.Entry(win,bd=5)
E1.grid(row=0,column=1)
E2=tkinter.Entry(win,bd=5)
E2.grid(row=1,column=1)

def inputfunc():

  path=(E1.get())
  path2=(E2.get())
  
  #path=input("enter path of chat doc")                             #enter the location of saved chat 
  path=path.strip('"' )
  #path2=input("enter path of excel sheet where you want to save")   #enter the excel sheet location where you want to save
  path2=path2.strip('"')

  chat=open(path,"r")    
  name=[]
  for i in chat:
    name.append(i)
  tokens=[]
  for i in name:
      try:
        a=i.split()[2]
        if not a.isalnum():
          tokens.append(a)
      except IndexError:
        continue
  data=[i.split('_') for i in tokens ]
  data = list(itertools.chain.from_iterable(data))    
  roll_no=[x for x in data if x.isdigit()]
  names=[x for x in data if x.isalpha()]
  
  if len(names)==len(roll_no):
    x={"ROLL NUMBER":roll_no,"Student Name":names}
    Attendance=pd.DataFrame(x)
  else:
    print("Invalid Data")

    
  Attendance= Attendance.drop_duplicates(subset=['ROLL NUMBER','Student Name'],keep='first')
  Attendance.to_excel(path2)                                           #Excel sheet saved
  mb.showinfo("Successful","Attendance Sheet Formed!")
  


def clear():
  E1.delete(0, tkinter.END)
  E2.delete(0, tkinter.END)
  

  

B1=tkinter.Button(win,text=" Exit ",command=win.destroy)
B1.grid(row=4,column=3)
B2=tkinter.Button(win,text="Generate",command=inputfunc)
B2.grid(row=4,column=1)
B3=tkinter.Button(win,text="New Path",command=clear)
B3.grid(row=4,column=0)
win.mainloop()

