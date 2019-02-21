from pandas import DataFrame
from pandas import read_excel
import os
import PySimpleGUI as sg

from requests import get

from os import path
from os import makedirs

from tkinter import filedialog
import tkinter as tk
import traceback



def processor(inputFilePath, outputDir):
    
    df = read_excel(inputFilePath)
    log = open("log.txt", "w")
    error = []
    
    for index, row in df.iterrows():
        try:
            path = row['Path']
            partNumber = row['Part Number']
            fileName = partNumber+'_'+ str(index)+'.pdf'
            directory = outputDir + '/' + row['Commodity']+'/'
            if not os.path.exists(directory):
                makedirs(directory)
            r=get(path)
            with open(directory+fileName,'wb') as f:
                f.write(r.content)
            print("creating DB COnn", file=log)
        except Exception as e:
            error.append(e)
            traceback.print_exc(file=log)
            pass
    if len(error) == 0:
        print("All files have been downloaded")
    else:
        print(error)
             
root = tk.Tk()
root.withdraw()
inputCSV =  filedialog.askopenfilename(initialdir = "/",title = "Select file",filetypes = (("Excel Files","*.xlsx"),("all files","*.*")))



root.withdraw()
output = filedialog.askdirectory(initialdir = "/", title = "Select Storage Folder")
root.destroy()


base = None

os.environ["TCL_LIBRARY"] = "C:\\Users\\vvandi\\AppData\\Local\\Continuum\\anaconda3\\pkgs\\python-3.7.1-h8c8aaf0_6\\tcl\\tcl8.6"
os.environ["TK_LIBRARY"] ="C:\\Users\\vvandi\\AppData\\Local\\Continuum\\anaconda3\\pkgs\\python-3.7.1-h8c8aaf0_6\\tcl\\tk8.6"

processor(inputCSV, output)