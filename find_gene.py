# -*- coding: utf-8 -*-
"""
Created on Tue Dec 15 11:05:41 2020

@author: td
"""
import tkinter
from tkinter import filedialog
import pandas as pd
import xlrd
import os
import openpyxl
from tkinter import*
import PySimpleGUI as SG
import re
from xlrd import XLRDError
  

def my_fun():
    if str(entery_2.get())=="":
       label_4 = Label(root, text="Error need input path!!!!!",width=50,font=("bold", 10),bg='red')
       label_4.place(x=50,y=330)
    elif str(entry_1.get())=="":
       label_4 = Label(root, text="Error need gene symbol!!!!!",width=50,font=("bold", 10),bg='red')
       label_4.place(x=50,y=330)
    elif str(entery_3.get())=="":
       label_4 = Label(root, text="Error need output path!!!!!",width=50,font=("bold", 10),bg='red')
       label_4.place(x=50,y=330)
    else:
       my_fun2()
       label_4 = Label(root, text="Search is completed!!!!! \n You can find your result at: \n"+ str(entery_3.get())+"/"+str(entry_1.get()).lower()+'_Search_result.csv',width=60,font=("bold", 10),bg='red')
       label_4.place(x=50,y=330)
        
def my_fun2():
        mydir =str(entery_2.get())
        mygene=str(entry_1.get())
        print(mygene)
        print(mydir)
        filelist=[]
        for path, subdirs, files in os.walk(mydir):
            for file in files:                
                if (file.endswith(str(var.get()))):# or file.endswith('.xlsx') or file.endswith('.xls') or file.endswith('.XLS')):
                    filelist.append(os.path.join(path, file))   
                    
        number_of_files=len(filelist)
        df_output = pd.DataFrame(pd.np.empty((number_of_files, 4)) * pd.np.nan,columns=['file_name','sheet','row','col'])
        count=0
        print(number_of_files)
        
        for i in range(number_of_files):               
            if (filelist[i].endswith('.txt')):
                str1=filelist[i]
                str1.replace('\\','/')
                with open(str1,'r') as file: 
                    num_line=0
                    for line in file: 
                        num_line=num_line+1
                        num_word=0
                        for word in line.split():  
                            num_word=num_word+1
                            if re.search(mygene.lower(), word.lower()):  
                                   df_output.loc[count,'file_name']=filelist[i]
                                   df_output.loc[count,'sheet']=""
                                   df_output.loc[count,'row']=num_line
                                   df_output.loc[count,'col']=num_word
                                   count=count+1 
                          
            else: 
                if (filelist[i].endswith('.xlsx') or filelist[i].endswith('.xls') or filelist[i].endswith('.XLS')):
                    str1=filelist[i]
                    str1.replace('\\','/')
                    try:
                        xls = pd.ExcelFile(str1)
                        for sheet in xls.sheet_names:                            
                            df1=[]
                            df1 = pd.read_excel(str1, sheet_name=sheet)
                            for row in range(len(df1)):
                                for col in range(len(df1.columns)):
                                    if re.search(mygene.lower(), str(df1.iloc[row,col]).lower()):# and df1.iloc[row,col]!="" :                                   
                                        df_output.loc[count,'file_name']=str1
                                        df_output.loc[count,'sheet']=sheet
                                        df_output.loc[count,'row']=row+2
                                        df_output.loc[count,'col']=col+1
                                        count=count+1 
                    except XLRDError:
                           label_4 = Label(root, text="Error: permission denied: "+str1+" ",width=20,font=("bold", 10))
                           label_4.place(x=100,y=330)
                                   
        df_output.to_csv(str(entery_3.get())+"/"+mygene.lower()+'_Search_result.csv')
        print(df_output)  

def fun_get_folder1():
    global folder_path
    filename = filedialog.askdirectory(title='Open folder to search')
    folder_path=filename
    print(folder_path)
    entery_2.delete(0, "end")
    entery_2.insert(END, str(folder_path))
    
def fun_get_folder2():
    global folder_path
    filename = filedialog.askdirectory(title='Open folder to search')
    folder_path=filename
    print(folder_path)
    entery_3.delete(0, "end")
    entery_3.insert(END, str(folder_path))    
  
def sel():
   selection = "You selected the option " + str(var.get())
   label.config(text = selection)
    
root = Tk()
folder_path = StringVar()
var = StringVar()
root.geometry('600x400')
root.title("Welcome to find gene")
label_0 = Label(root, text="Welcome to find gene",width=20,font=("bold", 20))
label_0.place(x=150,y=50)
label_1 = Label(root, text="Gene symbol:",width=20,font=("bold", 10))
label_1.place(x=50,y=250)
entry_1 = Entry(root)
entry_1.place(x=200,y=250)
label_2 = Label(root, text="Input path:",width=20,font=("bold", 10))
label_2.place(x=50,y=100)
entery_2= Entry(root)
entery_2.place(x=200,y=100)
Button1=Button(root, text='Browse',width=20,bg='black',fg='white',command =fun_get_folder1).place(x=350,y=100)
label_3 = Label(root, text="Output path:",width=20,font=("bold", 10))
label_3.place(x=50,y=150)
entery_3= Entry(root)
entery_3.place(x=200,y=150)
Button2=Button(root, text='Browse',width=20,bg='black',fg='white',command =fun_get_folder2).place(x=350,y=150)

label_4 = Label(root, text="File type:",width=20,font=("bold", 10))
label_4.place(x=50,y=200)

R1 = Radiobutton(root, text=".txt",  value=".txt", command=sel)
R1.place(x=200,y=200)

R2 = Radiobutton(root, text=".xlsx", value=".xlsx", command=sel)
R2.place(x=300,y=200)

R3 = Radiobutton(root, text=".xls", value=".xls", command=sel)
R3.place(x=400,y=200)

label = Label(root)

Button3=Button(root, text='Submit',width=20,bg='brown',fg='white',command = my_fun).place(x=350,y=250)
# it is use for display the registration form on the window
root.mainloop()



        
