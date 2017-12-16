#------------------
#TaxScraper
#v 1.0.0
#------------------
#Patrick Leahey
#For
#Nesreen Qanah
#12/6/17
#------------------


#------------------------------------------------------------------------------------------------
#                                         Information:

#Python 2.7.10

#Uses Modules
#   Selenium
#   Numpy
#   Pandas
#   Tkinter

#Must download:
#   PhantomJS - webdriver
#   Download at:
#       http://phantomjs.org/download.html
#   Unzip and place in directory
#   Change path to match location of phantomjs.exe (in bin directory in phantomjs version folder)
#------------------------------------------------------------------------------------------------

#-------------------------
#Import external libraries
#-------------------------
from selenium import webdriver
import numpy
import pandas as pd
from pandas import ExcelWriter
from pandas import ExcelFile
from Tkinter import *
import tkFileDialog

print(numpy.__file__)
print(pd.__file__)
print(tkFileDialog.__file__)


# Web Driver $$$$$$ MUST CHANGE TO YOUR OWN PATHWAY $$$$$$
driver = webdriver.PhantomJS(executable_path='/Users/Patrick/Desktop/phantomjs-2.1.1-macosx/bin/phantomjs')

#-------------------------
#County Specic Webscraping
#-------------------------
def ALAMANCENC(driver,url):
    driver.get(url)
    infoList = []
    for i in range(2,20,2):
        owed = driver.find_element_by_css_selector('body > table:nth-child(7) > tbody > tr > td > table.a1 > tbody > tr:nth-child(' + str(i) + ') > td:nth-child(8)').text
        year = driver.find_element_by_css_selector('body > table:nth-child(7) > tbody > tr > td > table.a1 > tbody > tr:nth-child(' + str(i) + ') > td:nth-child(2)').text
        infoList.append([owed == '$ 0.00',owed,year])
    return infoList

#---------------------------
#Get user-selected save path
#---------------------------
def getSavePath():
    path = tkFileDialog.asksaveasfile(defaultextension = '.xlsx')
    return path

#------
#Export
#------
def list_to_df(list):
    final_df = pd.DataFrame(list)
    #final_df.index.name = 'Index'
    #final_df.reset_index(inplace=True)
    final_df.columns = ['PARNO', 'OWNNAME', 'LOCSTATE', 'LOCCOUNTY', 'TAX YEAR', 'STATUS', 'Owed']
    path = getSavePath()
    writer = pd.ExcelWriter(path, engine='xlsxwriter')
    final_df.to_excel(writer, sheet_name='Sheet1')
    writer.save()

#-----------------------------
#General processing of queries
#-----------------------------
def process(path, driver, root,excelLabel,browseButton):

    #Get rid of old Tkinter widgets while retaining root
    excelLabel.destroy()
    browseButton.destroy()

    #Dictionary to hold all gathered information
    #Of the form {TAX ID: Owner Name, State, County, Year, Paid/Delinquent, Owed}
    allList = []

    #Load Data Frame
    pd.set_option('display.precision', 2)
    df = pd.read_excel(path)

    #Call county function for each property
    for i in df.index:
        #Provide progress update for user
        progressUpdate = Label(root,text= str((i+1)/len(df.index)*100) + '% \n Queries Complete', bg='light blue', fg='Dark Red', height=100, width=100)
        progressUpdate.config(font = ('helvetica',50))
        progressUpdate.place(relx=0.5, rely=0.5, anchor=CENTER)

        #Inspect df
        url = df['Tax collector Parcel link'][i]
        county = df['LOCCOUNTY'][i]
        state = df['LOCSTATE'][i]
        name = df['OWNNAME'][i]

        #Update dict
        if county == 'ALAMANCE' and state == 'NC':
            owedList = ALAMANCENC(driver,url)
            for shortList in owedList:
                if shortList[0]:
                    allList.append([df['PARNO'][i],name,state, county, shortList[2], 'Paid', shortList[1]])
                    #dict[df['PARNO'][i]] = (name,state, county, 2017, 'Paid', shortList[1])
                else:
                    allList.append([df['PARNO'][i], name, state, county, shortList[2], 'Delinquent', shortList[1]])
                    #dict[df['PARNO'][i]] = (name, state, county, 2017, 'Delinquent', shortList[1])

    #Call list_to_df, passing allList
    list_to_df(allList)

#---------------------
#Driver for GUI (main)
#---------------------
def main(driver):
    #Establish root
    root = Tk()
    root.wm_title('Property Tax Lookup Assistant')
    root.resizable(width=False, height=False)
    root.geometry('600x400')
    root.configure(background = 'light blue')

    #Title Label
    titleLabel= Label(root, text='TaxScraper', bg='light blue', fg='Dark Red', height=5, width=60)
    titleLabel.config(font=('helvetica', 40))
    titleLabel.place(relx=0.5, rely=0.25, anchor=CENTER)

    #Excel Label
    excelLabel = Label(root, text='Instructions: \n --------------------- \n Upload a .xlsx file with Urls, State, County, and Owner Name in columns named: \n "Tax collector Parcel link" \n "LOCSTATE" \n "LOCCOUNTY" \n "OWNNAME"', bg='light blue', fg='black', height=8, width=60)
    excelLabel.place(relx=0.5, rely=0.55, anchor=CENTER)

    #Patrick Leahey Label
    pLabel = Label(root, text='TaxScraper v1.0.0: Patrick Leahey',bg='light blue', fg='black', height=1, width=60)
    pLabel.place(relx=0.81, rely=0.97, anchor=CENTER)

    #Opens window that allows user to select path
    def getOpenPath():
        path = tkFileDialog.askopenfilename()
        process(path, driver, root, excelLabel, browseButton)

    #Browse Button
    browseButton = Button(root, text='Browse', command=getOpenPath, width=30, bg = 'light gray')
    browseButton.place(relx=0.5, rely=0.8, anchor=CENTER)
    root.mainloop()

main(driver)
driver.quit()
