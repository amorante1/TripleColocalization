import sys
from PIL import ImageEnhance, ImageFont, ImageDraw
import PIL.Image
import os
from os import listdir
from os.path import isfile, join
import xlwt
from xlwt import Workbook
from xlrd import open_workbook
from xlutils.copy import copy
from tkinter import *
from tkinter import filedialog
from tkinter import messagebox
import math


class Triple:

    def __init__(self):
        #initializes GUI.
        self.root = Tk()
        self.root.title('Triple Colocalization Analysis')
        self.root.geometry("1000x700")
        self.innerBox = Frame(self.root, width=1000, height=700, bg='white')
        self.innerBox.pack()
        #prompts user for folder containing images to analyze.
        self.mainFold = filedialog.askdirectory(title = 'Please Select  Folder with Three Color Channels')
        os.chdir(self.mainFold)
        self.onlyfiles = []
        self.length = 0
        self.newFold = str(self.mainFold) + '\Colocalization'
        #checks for folder presence and three image files. Shuts down if fails.
        if self.fileCheck() == 0:
            sys.exit()
        #converts images to grayscale
        self.resizeImages()
        #initializes excel sheet
        self.wb = Workbook()
        self.sheet1 = self.wb.add_sheet('Sheet 1')
        self.row = 1
        self.excelCheck()
        #starts main window.
        self.previewScreen()
        self.root.mainloop()

    def excelCheck(self):
        #checks whether excel file is present.
        os.chdir(self.newFold)
        #If excel file is present, it will start to add data under old data.
        if os.path.isfile('Colocalization Data.xls'):
            self.rb = open_workbook('Colocalization Data.xls')
            firstSheet = self.rb.sheet_by_index(0)
            self.placeHold = 1
            self.rowCheck(firstSheet)
            self.wb = copy(self.rb)
            self.sheet1 = self.wb.get_sheet(0)
        #if no excel sheet is found, it will make a new one.
        else:
            self.wb = Workbook()
            self.sheet1 = self.wb.add_sheet('Sheet 1')
            style = xlwt.easyxf('font: bold 1')
            self.sheet1.write(0, 1, 'mM1 1 to 2 and 3', style)
            self.sheet1.write(0, 2, 'mM2 2 to 1 and 3', style)
            self.sheet1.write(0, 3, 'mM3 3 to 1 and 2', style)
            self.sheet1.write(0, 4, 'Pearson Coefficient', style)
        os.chdir(self.mainFold)

    def rowCheck(self, sheet):
        #checks for lowest row not filled with data in excel sheet.
        while self.placeHold == 1:
            try:
                val = sheet.cell_value(rowx=self.row, colx=1)
            except IndexError:
                self.placeHold = 2
                return self.rowCheck(sheet)
            self.row +=1

    def fileCheck(self):
        # gets all files in main folder.
        self.onlyfiles = [f for f in listdir(self.mainFold) if os.path.isfile(join(self.mainFold, f))]
        #checks for Colocalization folder. Makes a new one if not present.
        if not os.path.exists(self.newFold):
            os.makedirs(self.newFold)
        self.length = len(self.onlyfiles)
        #checks to see if there are only three files present.
        if self.length != 3:
            messagebox.showerror('Error', 'Error. Please only have three image files in selected folder.')
            return 0
        else:
            return 1

    def resizeImages(self):
        #converts images to grayscale. Do not mind that function does not actually resize image as name suggests.
        j = 1
        for i in self.onlyfiles:
            currPic = PIL.Image.open(i)
            currPic = currPic.convert('L')
            os.chdir(self.newFold)
            currPic.save('Channel' + str(j) + 'Gray.png')
            os.chdir(self.mainFold)
            j += 1

    def previewScreen(self):
        #sets up main GUI.
        os.chdir(self.newFold)
        #places first color channel into GUI
        self.img1 = PIL.Image.open('Channel1Gray.png').resize((250,250))
        self.img1.save("C1G500.png")
        self.img1 = PhotoImage(file='C1G500.png')
        self.img1PH = Label(self.innerBox, image=self.img1)
        self.img1PH.image = self.img1
        self.img1PH.grid(row=1, column=1)
        #places second color channel into GUI
        self.img2 = PIL.Image.open('Channel2Gray.png').resize((250, 250))
        self.img2.save("C2G500.png")
        self.img2 = PhotoImage(file='C2G500.png')
        self.img2PH = Label(self.innerBox, image=self.img2)
        self.img2PH.image = self.img2
        self.img2PH.grid(row=1, column=2)
        #places third color channel into GUI
        self.img3 = PIL.Image.open('Channel3Gray.png').resize((250, 250))
        self.img3.save("C3G500.png")
        self.img3 = PhotoImage(file='C3G500.png')
        self.img3PH = Label(self.innerBox, image=self.img3)
        self.img3PH.image = self.img3
        self.img3PH.grid(row=1, column=3)
        #labels brightness scale
        self.BrightLabel = Label(self.innerBox, text = 'Brightness:', bg= 'white').grid(row=2, column = 0)
        #creates brightness scalebar for each color channel.
        self.Scale1 = Scale(self.innerBox, from_=-9, to=10, orient= HORIZONTAL)
        self.Scale1.grid(row = 2, column = 1, sticky=NSEW)
        self.Scale2 = Scale(self.innerBox, from_=-9, to=10, orient= HORIZONTAL)
        self.Scale2.grid(row=2, column=2, sticky=NSEW)
        self.Scale3 = Scale(self.innerBox, from_=-9, to=10, orient= HORIZONTAL)
        self.Scale3.grid(row=2, column=3, sticky=NSEW)
        self.Scale1.set(0)
        self.Scale2.set(0)
        self.Scale3.set(0)
        self.ContLabel = Label(self.innerBox, text='Contrast:', bg='white').grid(row=3, column=0)
        self.Scale4 = Scale(self.innerBox, from_=-10, to=10, orient=HORIZONTAL)
        self.Scale4.grid(row=3, column=1, sticky=NSEW)
        self.Scale5 = Scale(self.innerBox, from_=-10, to=10, orient=HORIZONTAL)
        self.Scale5.grid(row=3, column=2, sticky=NSEW)
        self.Scale6 = Scale(self.innerBox, from_=-10, to=10, orient=HORIZONTAL)
        self.Scale6.grid(row=3, column=3, sticky=NSEW)
        #self.Scale4.set(0)
        self.Scale5.set(0)
        self.Scale6.set(0)
        #creates button to update brightness settings
        self.updateButton = Button(self.innerBox, text='Update', command=self.setThreshold)
        self.updateButton.grid(row = 4, column = 2)
        #creates button to start colocalization analysis.
        self.AnalyButton = Button(self.innerBox, text='Analyze', command=self.colocAnalysis)
        self.AnalyButton.grid(row=5, column=2)
        self.root.update_idletasks()
        self.root.update()

    def setThreshold(self):
        #gets scalebar data for each color channel
        b1 = float(self.Scale1.get() + 2)/2
        b2 = float(self.Scale2.get() + 2) / 2
        b3 = float(self.Scale3.get() + 2) / 2
        #adjusts negative numbers and converts to decimals.
        if b1 <= 1:
            b1 = 1 + ((b1*2)-2)/10
        if b2 <= 1:
            b2 = 1 + ((b2*2)-2)/10
        if b3 <= 1:
            b3 = 1 + ((b3*2)-2)/10
        con1 = float(self.Scale4.get() + 2) / 2
        con2 = float(self.Scale5.get() + 2) / 2
        con3 = float(self.Scale6.get() + 2) / 2
        if con1 <= 1:
            con1 = 1 + ((con1 * 2) - 2) / 10
        else:
            con1 = con1 * 1
        if con2 <= 1:
            con2 = 1 + ((con2 * 2) - 2) / 10
        else:
            con2 = con2 * 1
        if con3 <= 1:
            con3 = 1 + ((con3 * 2) - 2) / 10
        else:
            con3 = con3 * 1
        #adjusts brightness for color channel 1.
        c1 = PIL.Image.open('Channel1Gray.png')
        enhancer = ImageEnhance.Brightness(c1)
        enhanceC = ImageEnhance.Contrast(c1)
        os.chdir(self.newFold)
        c1 = enhanceC.enhance(con1).save('Channel1Gray.png')
        c1 = enhancer.enhance(b1).save('Channel1Gray.png')
        # adjusts brightness for color channel 2.
        c1 = PIL.Image.open('Channel2Gray.png')
        enhancer = ImageEnhance.Brightness(c1)
        enhanceC = ImageEnhance.Contrast(c1)
        c1 = enhanceC.enhance(con2).save('Channel2Gray.png')
        c1 = enhancer.enhance(b2).save('Channel2Gray.png')
        # adjusts brightness for color channel 3.
        c1 = PIL.Image.open('Channel3Gray.png')
        enhancer = ImageEnhance.Brightness(c1)
        enhanceC = ImageEnhance.Contrast(c1)
        c1 = enhanceC.enhance(con3).save('Channel3Gray.png')
        c1 = enhancer.enhance(b3).save('Channel3Gray.png')
        #destroys old preview screen and makes a new one with the updated pictures
        for child in self.innerBox.winfo_children():
            child.destroy()
        self.previewScreen()

    def colocAnalysis(self):
        #collects pixel data for all images and places into separate lists.
        image = PIL.Image.open('Channel1Gray.png')
        data1 = image.getdata()
        image = PIL.Image.open('Channel2Gray.png')
        data2 = image.getdata()
        image = PIL.Image.open('Channel3Gray.png')
        data3 = image.getdata()
        #checks that each image contains same amount of pixels. Shuts down if fails.
        if len(data1) != len(data2) or len(data2) != len(data3) or len(data1) != len(data3):
            messagebox.showerror('Error', 'Error. Please make sure images are the same size.')
            os.remove('Channel1Gray.png')
            os.remove('Channel2Gray.png')
            os.remove('Channel3Gray.png')
            sys.exit()
        totalif = 0
        totals = 0
        #calculates Manders' coefficient for channel 1 to channels 2 and 3.
        #Manders' coefficient = sum of C1 intensities overlapping with C2 AND C3 divided by sum of all C1 intensities.
        for i in range(len(data1)):
            if data2[i] > 1 and data3[i] > 1:
                totalif += data1[i]
            if data1[i] > 1:
                totals += data1[i]
        m1 = totalif / totals
        totalif = 0
        totals = 1
        self.sheet1.write(self.row, 1, m1)
        # calculates Manders' coefficient for channel 2 to channels 1 and 3.
        # Manders' coefficient = sum of C2 intensities overlapping with C1 AND C3 divided by sum of all C1 intensities.
        for i in range(len(data2)):
            if data1[i] > 1 and data3[i] > 1:
                totalif += data2[i]
            if data2[i] > 1:
                totals += data2[i]
        m1 = totalif / totals
        totalif = 0
        totals = 1
        self.sheet1.write(self.row, 2, m1)
        # calculates Manders' coefficient for channel 3 to channels 1 and 2.
        # Manders' coefficient = sum of C3 intensities overlapping with C1 AND C2 divided by sum of all C3 intensities.
        for i in range(len(data3)):
            if data2[i] > 1 and data1[i] > 1:
                totalif += data3[i]
            if data3[i] > 1:
                totals += data3[i]
        m1 = totalif / totals
        totalif = 0
        totals = 1
        self.sheet1.write(self.row, 3, m1)
        #apply pearson coefficient formula to data by comparing pixel values from channel 1 to the merge of channels 2 and 3.
        data5 = []
        for i in range(len(data1)):
            if data2[i] > 5 and data3[i] > 5:
                data5.append(data2[i])
            else:
                data5.append(0)
        avg2 = sum(data5) / len(data5)
        avg3 = sum(data1) / len(data1)
        pTotal = 0
        sumRsq = 0
        sumGsq = 0
        for i in range(len(data5)):
            pTotal += ((data5[i] - avg2) * (data1[i] - avg3))
            sumRsq += (data5[i] - avg2) ** 2
            sumGsq += (data1[i] - avg3) ** 2
        pears = pTotal / ((sumRsq * sumGsq) ** 0.5)
        self.sheet1.write(self.row, 4, pears)
        os.chdir(self.newFold)
        #saves data into excel sheet and cleans up.
        self.wb.save('Colocalization Data.xls')
        os.remove('Channel1Gray.png')
        os.remove('Channel2Gray.png')
        os.remove('Channel3Gray.png')
        os.remove('C1G500.png')
        os.remove('C2G500.png')
        os.remove('C3G500.png')
        sys.exit()

Triple()









