import pandas as pd
import tkinter as userinterace
import pyttsx3
import PyPDF2
import fpdf
import filetype
from fpdf import FPDF
import os
import docx2pdf
from docx2pdf import convert
from tkinter import messagebox


# Press the green button in the gutter to run the script.

def readAgent(tempbook ,begin):
    # Reading process to begin
    num=int(begin)
    speakAgent = pyttsx3.init()
    speakAgent.getProperty('rate')
    speakAgent.setProperty('rate', 150)
    speakAgent.getProperty('volume')
    speakAgent.setProperty('volume', 0.8)
    book = open(tempbook, 'rb')
    pdfreader = PyPDF2.PdfFileReader(book)
    totalPages = pdfreader.numPages
    if(totalPages==1):
        pdfPage = pdfreader.getPage(0)
        text += pdfPage.extractText().split()
        print(text)
        speakAgent.say(text)
        speakAgent.runAndWait()
    else:
     for i in range(num-1, totalPages):
        print(i)
        pdfPage = pdfreader.getPage(i)
        text = pdfPage.extractText()
        print(text)
        speakAgent.say(text)
        speakAgent.runAndWait()

def readFileCheck(loc,begin):
    begin=begin
    ext = os.path.splitext(loc)[-1].lower()
    print(ext)
    trial=filetype.guess(loc)
    print(trial)
    if (ext == '.pdf'):
         readAgent(loc,begin)
    elif(ext == '.txt'):
        print("Converting the .txt file ")
        pdf = FPDF()
        pdf.add_page()
        pdf.set_font("Arial", size=15)
        f = open(loc, "r")
        for x in f:
            pdf.cell(200, 10, txt= x, ln=1, align='C')
        pdf.output("C:\\Users\\subha\\OneDrive\\Desktop\\temp.pdf")
        print("File converted")
        path="C:\\Users\\subha\\OneDrive\\Desktop\\temp.pdf"
        readAgent(path,begin)
    elif(ext=='.doc' or ext =='.docx'):
        basefile=os.path.basename(loc)
        path = "C:\\Users\\subha\\OneDrive\\Desktop\\temp.pdf"
        convert(loc,path)
        readAgent(path, begin)



def speakerInitiate():

    fileloc=inputPath.get() #Location of the file to read
    pageNum=inputCount.get() # Page number to start with it

    # checking if a location is correct or not
    if(len(fileloc)== 0):
        message1=userinterace.Toplevel()
        label4=userinterace.Label(master=message1,height=10,width=40,text="Error Found: Need to give the file location")
        label4.config(font=('Verdana 20 bold',15))
        label4.pack()
        message1.mainloop()
        exit()
    else:
        # Checking if a given page number is correct
        if(len(pageNum)==0 or int(pageNum)==0):
            messagebox.showerror("Error Message", "Invalid page number. Please check")
            exit()
        else:
         #short message of the speaker
         speakAgent = pyttsx3.init()
         speakAgent.say("Your audio book will start  in a short while..")
         speakAgent.runAndWait()
         readFileCheck(fileloc,pageNum)


if __name__ == '__main__':

   # tkinter initiation
    root = userinterace.Tk()
   # simple GUI design
    canvas1 = userinterace.Canvas(master=root, width=800, height=300, bg="light yellow")
    canvas1.pack()
    titlelabel = userinterace.Label(master=root, text='PYTHON PDF READER', bg="light yellow")
    titlelabel.config(font=('Verdana 20 bold', 15))
    canvas1.create_window(400, 30, window=titlelabel)

    label1 = userinterace.Label(master=root, text='Enter the file path',bg="light yellow")
    label1.config(font=('Arial', 15))
    canvas1.create_window(400,70, window=label1)

    inputPath = userinterace.Entry(root)
    canvas1.create_window(400,120, window=inputPath)

    label2 = userinterace.Label(master=canvas1, text='Enter the start page', bg="light yellow")
    label2.config(font=('Arial', 15))
    canvas1.create_window(400, 170, window=label2)

    inputCount = userinterace.Entry(root)
    canvas1.create_window(400, 220, window=inputCount)

    button1 = userinterace.Button(master=root, text='Read',bg='light green',font=('Arial', 11, 'bold'),command=speakerInitiate)
    canvas1.create_window(400, 250, window=button1)
    button1.pack()


    root.mainloop()

