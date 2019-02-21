#!/usr/bin/env python
# coding: utf-8

import requests
import pandas as pd 
import os
from tkinter import filedialog
import tkinter as tk
import sys

base = None

os.environ["TCL_LIBRARY"] = "C:\\Users\\vvandi\\AppData\\Local\\Continuum\\anaconda3\\pkgs\\python-3.7.1-h8c8aaf0_6\\tcl\\tcl8.6"
os.environ["TK_LIBRARY"] ="C:\\Users\\vvandi\\AppData\\Local\\Continuum\\anaconda3\\pkgs\\python-3.7.1-h8c8aaf0_6\\tcl\\tk8.6"


if sys.platform=='win32':
    base = "Win32GUI"



def processor(inputFilePath, outputDir):
    
    df = pd.read_excel(inputFilePath)
    
    for index, row in df.iterrows():
    
        path = row['Path']
        partNumber = row['Partnumber2']
        fileName = partNumber+'_'+ str(index)+'.pdf'
        directory = outputDir + '/' + row['Commodity']+'/'
        if not os.path.exists(directory):
            os.makedirs(directory)
        r=requests.get(path)
        with open(directory+fileName,'wb') as f:
            f.write(r.content)


root = tk.Tk()
tk.mainloop()
root.withdraw()
inputCSV =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel Files","*.xlsx"),("all files","*.*")))



root.withdraw()
output = filedialog.askdirectory(initialdir = "/", title = "Select Storage Folder")
root.destroy()



processor(inputCSV, output)