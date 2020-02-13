# -*- coding: utf-8 -*-
"""
Created on Fri Apr 12 15:56:34 2019

@author: adam.rees

This script will merge all pdfs within a folder and save a combined version to
the desktop.  This may be used for printing.
"""

import sys
import os
import winshell
import time
import glob
import tkinter
from tkinter import messagebox
import PyPDF4

start = time.time()

ClickedOnFolder = " ".join(sys.argv[1:])

Desktop = winshell.desktop()

def PDFSearch(folder):
    """
    This function searches for PDFs within the folder
    and returns a list of the pdf titles
    It expects the variable "folder"
    """
    os.chdir(folder)
    PDFList=[]

    for pdf in glob.glob("*.pdf"):
        PDFList.append(pdf)
    
    CheckList(PDFList)

    return PDFList

def CheckList(List):
    """
    This function checks the length of a list, and if 0 it will
    send out a message box explaining that there are none found.
    It expects the variable "List"
    """
    NumberOfItems = len(List)

    if NumberOfItems == 0:
        root = tkinter.Tk()
        root.withdraw()
        messagebox.showinfo("Warning","No PDF files found in folder")
        sys.exit(4)
    else:
        pass

def PDFMergeFunc(List):
    pdfWriter = PyPDF4.PdfFileWriter()
    for pdf in List:
        pdfFileObj = open(pdf,'rb')
        pdfReader = PyPDF4.PdfFileReader(pdfFileObj)

        #Opening each page of the PDF
        for pageNum in range(pdfReader.numPages):
            pageObj = pdfReader.getPage(pageNum)
            pdfWriter.addPage(pageObj)

        #save PDF to file on desktop, wb for write binary
        pdfOutput = open(f"{os.path.join(Desktop,'PrintMe.pdf')}"
                            , 'wb')

        #Outputting the PDF & Closing the PDF writer
        pdfWriter.write(pdfOutput)
        pdfOutput.close()


PDFs = PDFSearch(ClickedOnFolder)
NumberOfPDFs = len(PDFs)
PDFMergeFunc(PDFs)
end = time.time()

print(f"""
      Done! Merged {NumberOfPDFs} PDF files in {round(end-start,2)} seconds.\n
      Processed at a rate of {round(round(end-start, 2)/NumberOfPDFs, 3)} 
      seconds per PDF.
      """)

time.sleep(2)
