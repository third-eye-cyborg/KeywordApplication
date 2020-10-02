from pytrends.request import TrendReq
import pandas as pd
from tkinter import *
import os

# create class
class MyKeywordApp():
    def __init__(self):
        # self.keyword = input('Enter Keyword:  ')
        self.newWindow()

    def newWindow(self):
        # define your window
        root = Tk()
        root.geometry("400x100")
        root.resizable(False, False)
        root.title("Keyword Application")

        # add logo image
        p1 = PhotoImage(file='logo_image.png')
        root.iconphoto(False, p1)

        # add labels
        label1 = Label(text='Input a Keyword')
        label1.pack()
        canvas1 = Canvas(root)
        canvas1.pack()
        entry1 = Entry(root)
        canvas1.create_window(200, 20, window=entry1)

        def excelWriter():
            # get the user-input variable
            x1 = entry1.get()
            canvas1.create_window(200, 210)

            # get our Google Trends data
            pytrend = TrendReq()
            kws = pytrend.suggestions(keyword=x1)
            df = pd.DataFrame(kws)
            df = df.drop(columns='mid')

            # create excel writer object
            writer = pd.ExcelWriter('keywords.xlsx')
            df.to_excel(writer)
            writer.save()

            # open your excel file
            os.system("keywords.xlsx")
            print(df)

        # add button and run loop
        button1 = Button(canvas1, text='Run Report', command=excelWriter)
        canvas1.create_window(200, 50, window=button1)
        root.mainloop()

# call class
MyKeywordApp()

