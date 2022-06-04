#Sales Tracker Frame

import tkinter as tk
import tkcalendar as tkc
import csv
import math as m
import calendar as c
from datetime import date
from datetime import datetime
from datetime import timedelta
import os.path
import win32ui
import os

##function definitions
def main():
    def dateToInteger(date):
        return int(10000*date.year + 100*date.month + date.day)

    #gets date of sunday
    def sunday():
        today = date.today()
        idx = (today.weekday() + 1) % 7
        sun = today - timedelta(idx)
        return sun

    #success return text
    def success():
        string = 'Your response has been submitted'
        return string

    #failure return text prompts
    def failure():
        string = 'There was an error, review submission'
        return string

    def name():
        string = 'Please choose a name'
        return string

    def attach():
        string = 'Please select attach options'
        return string

    def clear():
        string = 'Entry fields cleared'
        return string

    #makes name file
    def makeNameFile():
        fileName = 'associateNames.txt'
        if os.path.isfile(fileName):
            pass
        else:
            outputData = 'Names (One Name Per Line)\nEnter Names in associateNames.txt'
            with open(fileName, 'w', newline='') as f:
                f.write(outputData)

    #make name list
    def nameList():
        fileName = 'associateNames.txt'
        NAMES = []
        with open(fileName, 'r') as f:
            next(f)
            for line in f.readlines():
                name = line.strip()
                NAMES.append((name))
    ##        reader = csv.reader(f, skipinitialspace=True)
    ##        if csv.Sniffer().has_header:
    ##            next(reader)
    ##        NAMES = list(reader)
        return NAMES

    #Makes file for printer sales
    def makePrinterFile():
        fileName = 'printerSales.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Name', 'Date', 'Price', 'Subtotal', 'ESP', 'Ink']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #Makes file for computer sales
    def makeComputerFile():
        fileName = 'computerSales.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Name', 'Date', 'Price', 'Subtotal', 'ESP', 'Virus Shield']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #Makes file for chair sales
    def makeChairFile():
        fileName = 'chairSales.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Name', 'Date', 'Price', 'Subtotal', 'ESP', 'Chair Mat']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #makes return file
    def makePrinterReturns():
        fileName = 'printerReturns.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Date']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #makes return file
    def makeComputerReturns():
        fileName = 'computerReturns.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Date']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #makes return file
    def makeChairReturns():
        fileName = 'chairReturns.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Date']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #makes morning reports file
    def makeMornReports():
        fileName = 'morningReports.csv'
        if  os.path.isfile(fileName):
            pass
        else:
            outputData = [['Date Int', 'Date', 'Conversion', 'DPT', 'ROV', 'Comp', '# of Transactions', 'ESP Units', 'ESP %', 'ESP $', 'ET $', 'More Accounts', 'Total SAT']]
            with open(fileName, 'w', newline = '') as f:
                writer = csv.writer(f)
                writer.writerows(outputData)

    #makes daily summary file
    def makeDaily():
        fileName = datetime.now()
        if os.path.isfile(fileName.strftime("%d %B %Y")+".txt"):
            pass
        else:
            outputData = str(date.today())
            with open(fileName.strftime("%d %B %Y")+".txt", 'w', newline='') as f:
                f.write(outputData)

    #print daily summary file
    def printDaily():
        fileName = datetime.now()
        os.startfile(fileName.strftime("%d %B %Y")+".txt", "print")

    #printer sales input
    def computerInput():
        fileName = 'computerSales.csv'
        outputData = []

        data = []
        name = comp_name.get()
        today = dateToInteger(date.today())
        price = float(comp_price.get())
        subtotal = float(comp_sub.get())
        esp = comp_esp.get()
        vs = comp_vs.get()
        data.append(name)
        data.append(today)
        data.append(price)
        data.append(subtotal)
        data.append(esp)
        data.append(vs)

        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

    #printer submit
    def compSubmit():
        if comp_name.get() == 'Choose':
            comp_response['text']=name()
        else:
            if comp_esp.get() == 'Choose':
                comp_response['text']=attach()
            else:
                if comp_vs.get() == 'Choose':
                    comp_response['text']=attach()
                else:
                    try:
                        computerInput()
                        compPrint()
                        compClear()
                        comp_response['text']=success()
                    except:
                        comp_response['text']=failure()
                        

    #computer return input
    def computerRetIn():
        fileName = 'computerReturns.csv'
        outputData = []

        data = []
        today = dateToInteger(date.today())
        data.append(today)
        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

        comp_response['text']=success()

    #clears inputs
    def compClear():
        comp_name.set('Choose')
        comp_price.delete(0, 999)
        comp_sub.delete(0, 999)
        comp_esp.set('Choose')
        comp_vs.set('Choose')

        comp_response['text']=clear()

    #puts sale in daily sum file
    def compPrint():
        fileName = datetime.now()
        with open(fileName.strftime("%d %B %Y")+".txt", 'a', newline='') as f:
            name = comp_name.get()
            price = comp_price.get()
            sub = comp_sub.get()
            esp = comp_esp.get()
            vs = comp_vs.get()
            text = '\n\nItem: PC   Name: ' + str(name) +  '\nPrice of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Virus Shield: ' + str(vs)
            f.write(text)

    #prints submission
##    def compPrint():
##        dc = win32ui.CreateDC()
##        dc.CreatePrinterDC()
##        dc.StartDoc('Comp Receipt')
##        dc.StartPage()
##        today = date.today()
##        name = comp_name.get()
##        price = comp_price.get()
##        sub = comp_sub.get()
##        esp = comp_esp.get()
##        vs = comp_vs.get()
##        text = 'Item: PC   Name: ' + str(name) + '   Date: ' + str(today) + '   Price of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Virus Shield: ' + str(vs)
##        dc.TextOut(100,100, text)
##        dc.MoveTo(100, 102)
##        dc.LineTo(200, 102)
##        dc.EndPage()
##        dc.EndDoc()

    #printer sales input
    def printerInput():
        fileName = 'printerSales.csv'
        outputData = []

        data = []
        name = print_name.get()
        today = dateToInteger(date.today())
        price = float(print_price.get())
        subtotal = float(print_sub.get())
        esp = print_esp.get()
        ink = print_ink.get()
        data.append(name)
        data.append(today)
        data.append(price)
        data.append(subtotal)
        data.append(esp)
        data.append(ink)

        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

    #printer submit
    def printSubmit():
        if print_name.get() == 'Choose':
            print_response['text']=name()
        else:
            if print_esp.get() == 'Choose':
                print_response['text']=attach()
            else:
                if print_ink.get() == 'Choose':
                    print_response['text']=attach()
                else:
                    try:
                        printerInput()
                        printPrint()
                        printClear()
                        print_response['text']=success()
                    except:
                        print_response['text']=failure()

    #printer return input
    def printerRetIn():
        fileName = 'printerReturns.csv'
        outputData = []

        data = []
        today = dateToInteger(date.today())
        data.append(today)
        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)
            
        print_response['text']=success()

    #clears input
    def printClear():
        print_name.set('Choose')
        print_price.delete(0, 999)
        print_sub.delete(0, 999)
        print_esp.set('Choose')
        print_ink.set('Choose')

        print_response['text']=clear()

    def printPrint():
        fileName = datetime.now()
        with open(fileName.strftime("%d %B %Y")+".txt", 'a', newline='') as f:
            name = print_name.get()
            price = print_price.get()
            sub = print_sub.get()
            esp = print_esp.get()
            ink = print_ink.get()
            text = '\n\nItem: Printer   Name: ' + str(name) +  '\nPrice of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Ink: ' + str(ink)
            f.write(text)

##    #prints submission
##    def printPrint():
##        dc = win32ui.CreateDC()
##        dc.CreatePrinterDC()
##        dc.StartDoc('Print Receipt')
##        dc.StartPage()
##        today = date.today()
##        name = print_name.get()
##        price = print_price.get()
##        sub = print_sub.get()
##        esp = print_esp.get()
##        vs = print_ink.get()
##        text = 'Item: Printer   Name: ' + str(name) + '   Date: ' + str(today) + '   Price of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Ink: ' + str(vs)
##        dc.TextOut(100,100, text)
##        dc.MoveTo(100, 102)
##        dc.LineTo(200, 102)
##        dc.EndPage()
##        dc.EndDoc()

    #chair sales input
    def chairInput():
        fileName = 'chairSales.csv'
        outputData = []

        data = []
        name = chair_name.get()
        today = dateToInteger(date.today())
        price = float(chair_price.get())
        subtotal = float(chair_sub.get())
        esp = chair_esp.get()
        mat = chair_mat.get()
        data.append(name)
        data.append(today)
        data.append(price)
        data.append(subtotal)
        data.append(esp)
        data.append(mat)

        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

    #printer submit
    def chairSubmit():
        if chair_name.get() == 'Choose':
            chair_response['text']=name()
        else:
            if chair_esp.get() == 'Choose':
                chair_response['text']=attach()
            else:
                if chair_mat.get() == 'Choose':
                    char_response['text']=attach()
                else:
                    try:
                        chairInput()
                        chairPrint()
                        chairClear()
                        chair_response['text']=success()
                    except:
                        chair_response['text']=failure()

    #chair return input
    def chairRetIn():
        fileName = 'chairReturns.csv'
        outputData = []

        data = []
        today = dateToInteger(date.today())
        data.append(today)
        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

        chair_response['text']=success()

    #clears inputs
    def chairClear():
        chair_name.set('Choose')
        chair_price.delete(0, 999)
        chair_sub.delete(0, 999)
        chair_esp.set('Choose')
        chair_mat.set('Choose')

        chair_response['text']=clear()

    def chairPrint():
        fileName = datetime.now()
        with open(fileName.strftime("%d %B %Y")+".txt", 'a', newline='') as f:
            name = chair_name.get()
            price = chair_price.get()
            sub = chair_sub.get()
            esp = chair_esp.get()
            mat = chair_mat.get()
            text = '\n\nItem: Chair   Name: ' + str(name) +  '\nPrice of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Chair Mat: ' + str(mat)
            f.write(text)

##    #prints submission
##    def chairPrint():
##        dc = win32ui.CreateDC()
##        dc.CreatePrinterDC()
##        dc.StartDoc('Print Receipt')
##        dc.StartPage()
##        today = date.today()
##        name = chair_name.get()
##        price = chair_price.get()
##        sub = chair_sub.get()
##        esp = chair_esp.get()
##        vs = chair_mat.get()
##        text = 'Item: Chair   Name: ' + str(name) + '   Date: ' + str(today) + '   Price of item: $' + str(price) + '   Subtotal: $' + str(sub) + '   ESP: ' + str(esp) + '   Chair Mat: ' + str(vs)
##        dc.TextOut(100,100, text)
##        dc.MoveTo(100, 102)
##        dc.LineTo(200, 102)
##        dc.EndPage()
##        dc.EndDoc()

    #morning reports input
    def mornReports():
        fileName = 'morningReports.csv'
        outputData = []
        data = []
        today = date.today()
        yest = today - timedelta(1)
        dateInt = dateToInteger(yest)
        conversion = float(morn_conv.get())
        dpt = float(morn_dpt.get())
        rov = float(morn_prod.get())
        comp = float(morn_comp.get())
        trans = float(morn_tran.get())
        espUnits = float(morn_espu.get())
        espPerc = float(morn_esp.get())
        espD = float(morn_espd.get())
        etD = float(morn_etd.get())
        moreAcc = float(morn_ma.get())
        csat = float(morn_nps.get())
        data.append(dateInt)
        data.append(yest)
        data.append(conversion)
        data.append(dpt)
        data.append(rov)
        data.append(comp)
        data.append(trans)
        data.append(espUnits)
        data.append(espPerc)
        data.append(espD)
        data.append(etD)
        data.append(moreAcc)
        data.append(csat)

        outputData.append(data)

        with open(fileName, 'a', newline = '') as f:
            writer = csv.writer(f)
            writer.writerows(outputData)

    #submits morning reports
    def mornSubmit():
        try:
            mornReports()
            mornClear()
            morn_response['text']=success()
        except:
            morn_response['text']=failure()

    #clears entries
    def mornClear():
        morn_conv.delete(0,999)
        morn_dpt.delete(0,999)
        morn_prod.delete(0,999)
        morn_comp.delete(0,999)
        morn_tran.delete(0,999)
        morn_espu.delete(0,999)
        morn_esp.delete(0,999)
        morn_espd.delete(0,999)
        morn_etd.delete(0,999)
        morn_ma.delete(0,999)
        morn_nps.delete(0,999)

        morn_response['text']=clear()

    #averages for summary
    def avgConv():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[2])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgDPT():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[3])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgROV():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[4])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgComp():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[5])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgTrans():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[6])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgESPU():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[7])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgESPD():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[9])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgET():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[10])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgMA():
        fileName = 'morningReports.csv'
        total = 0
        count = 0
        today = date.today()
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(fileName) as f:
            reader = csv.reader(f)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[0]) >= startDate and int(row[0]) <= endDate:
                    total += float(row[11])
                    count += 1
        if count == 0:
            return 'No Reports'
        else:
            avg = round(total / count,2)
            return avg

    def avgPC():
        sales = 'computerSales.csv'
        returns = 'computerReturns.csv'
        subtotal = 0
        total = 0
        count = 0
        today = date.today()
        ifret = 0
        scount = 0
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    subtotal += float(row[2])
                    total += float(row[3])
                    count += 1
                    scount = count 
        if count <= 0:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
                    ifret = 1 
            if ifret == 1:
                avg = round((total-subtotal) / 1,2)
                string = str(avg) + '   ' + str(scount) + ' sold'
                return string
            else:
                return 'No Sales'
        else:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
            avg = round((total-subtotal) / count,2)
            string = str(avg) + '   ' + str(scount) + ' sold'
            return string

    def avgPrint():
        sales = 'printerSales.csv'
        returns = 'printerReturns.csv'
        subtotal = 0
        total = 0
        count = 0
        today = date.today()
        ifret = 0
        scount = 0
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    subtotal += float(row[2])
                    total += float(row[3])
                    count += 1
                    scount = count
        if count <= 0:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
                    ifret = 1 
            if ifret == 1:
                avg = round((total-subtotal) / 1,2)
                string = str(avg) + '   ' + str(scount) + ' sold'
                return string
            else:
                return 'No Sales'
        else:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
            avg = round((total-subtotal) / count,2)
            string = str(avg) + '   ' + str(scount) + ' sold'
            return string

    def avgChair():
        sales = 'chairSales.csv'
        returns = 'chairReturns.csv'
        subtotal = 0
        total = 0
        count = 0
        today = date.today()
        ifret = 0
        scount = 0
        if wtd_pw.get() == 'WTD':
            startDate = int(dateToInteger(sunday()))
            endDate = int(dateToInteger(today))
        elif wtd_pw.get() == 'PW':
            startDate = int(dateToInteger(sunday() - timedelta(7)))
            endDate = int(dateToInteger(today - timedelta(7)))
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    subtotal += float(row[2])
                    total += float(row[3])
                    count += 1
                    scount = count
        if count <= 0:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
                    ifret = 1 
            if ifret == 1:
                avg = round((total-subtotal) / 1,2)
                string = str(avg) + '   ' + str(scount) + ' sold'
                return string
            else:
                return 'No Sales'
        else:
            with open(returns) as r:
                reader = csv.reader(r)
                if csv.Sniffer().has_header:
                    next(reader)
                for row in reader:
                    if int(row[0]) >= startDate and int(row[0]) <= endDate:
                        count -= 1
            avg = round((total-subtotal) / count,2)
            string = str(avg) + '   ' + str(scount) + ' sold'
            return string

    def sumUpdate():
        sum_conv['text'] = avgConv()
        sum_dpt['text'] = avgDPT()
        sum_prod['text'] = avgROV()
        sum_comp['text'] = avgComp()
        sum_tran['text'] = avgTrans()
        sum_espu['text'] = avgESPU()
        sum_espd['text'] = avgESPD()
        sum_etd['text'] = avgET()
        sum_ma['text'] = avgMA()
        sum_pc['text'] = avgPC()
        sum_print['text'] = avgPrint()
        sum_chair['text'] = avgChair()

    #Lookup funcitons

    #name for lookup
    def lookName():
        name = look_name.get()
        return name

    #PC Basket
    def lookCompBask(name):
        sales = 'computerSales.csv'
        subtotal = 0
        total = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        subtotal += float(row[2])
                        total += float(row[3])
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((total-subtotal) / count,2)
            text = str(avg) + '  ' + str(count) + ' sold'
            return text

    #PC ESP
    def lookCompESP(name):
        sales = 'computerSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[4] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    #PC other attach
    def lookCompAtt(name):
        sales = 'computerSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[5] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    #Printer Basket
    def lookPrintBask(name):
        sales = 'printerSales.csv'
        subtotal = 0
        total = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        subtotal += float(row[2])
                        total += float(row[3])
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((total-subtotal) / count,2)
            text = str(avg) + '  ' + str(count) + ' sold'
            return text

    #Printer ESP
    def lookPrintESP(name):
        sales = 'printerSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[4] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    def lookPrintAtt(name):
        sales = 'printerSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[5] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    #Chair Basket
    def lookChairBask(name):
        sales = 'chairSales.csv'
        subtotal = 0
        total = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        subtotal += float(row[2])
                        total += float(row[3])
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((total-subtotal) / count,2)
            text = str(avg) + '  ' + str(count) + ' sold'
            return text

    #Printer ESP
    def lookChairESP(name):
        sales = 'chairSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[4] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    def lookChairAtt(name):
        sales = 'chairSales.csv'
        y = 0
        count = 0
        startDate = dateToInteger(look_start_date.get_date())
        endDate = dateToInteger(look_end_date.get_date())
        with open(sales) as s:
            reader = csv.reader(s)
            if csv.Sniffer().has_header:
                next(reader)
            for row in reader:
                if int(row[1]) >= startDate and int(row[1]) <= endDate:
                    if row[0] == name:
                        if row[5] == 'Yes':
                            y += 1
                        count += 1
        if count == 0:
            return 'No Sales'
        else:
            avg = round((y/count)*100,2)
            return avg

    #look search
    def lookSearch(name):
        if look_name.get() == 'Choose':
            look_comp_basket['text']='Enter Name'
        elif dateToInteger(look_start_date.get_date()) > dateToInteger(look_end_date.get_date()):
            look_comp_basket['text'] = 'Check Dates'
        else:
            name = lookName()
            look_comp_basket['text'] = lookCompBask(name)
            look_comp_esp['text'] = lookCompESP(name)
            look_comp_vs['text'] = lookCompAtt(name)
            look_print_basket['text'] = lookPrintBask(name)
            look_print_esp['text'] = lookPrintESP(name)
            look_print_ink['text'] = lookPrintAtt(name)
            look_chair_basket['text'] = lookChairBask(name)
            look_chair_esp['text'] = lookChairESP(name)
            look_chair_mat['text'] = lookChairAtt(name)

    #lookup print all
    def printAll():
        fileName = 'All Report.txt'
        NAMES = nameList()
        startDate = look_start_date.get_date()
        endDate = look_end_date.get_date()
        with open(fileName, 'w') as f:
            f.write(str(startDate) + ' - ' + str(endDate))
            f.write('\n')
            for name in NAMES:
                f.write(name)
                f.write('\nPC Basket: ' + str(lookCompBask(name)) + '  ESP Attach: ' + str(lookCompESP(name)) + '  Virus Shield Attach: ' + str(lookCompAtt(name)))
                f.write('\nPrinter Basket: ' + str(lookPrintBask(name)) + '  ESP Attach: ' + str(lookPrintESP(name)) + '  Ink/Toner Attach: ' + str(lookPrintAtt(name)))
                f.write('\nChair Basket: ' + str(lookChairBask(name)) + '  ESP Attach: ' + str(lookChairESP(name)) + '  Chair Mat Attach: ' + str(lookChairAtt(name)))
                f.write('\n\n\n')
        os.startfile(fileName, "print")
        


    #generates required data files if not present
    makeNameFile()
    makeComputerFile()
    makePrinterFile()
    makeChairFile()
    makeComputerReturns()
    makePrinterReturns()
    makeChairReturns()
    makeMornReports()
    makeDaily()


    ###
    ### THIS IS THE START OF THE GUI INTERFACE
    ###


    # Set dimensions
    HEIGHT = 800
    WIDTH = 1400

    # Fonts
    HEADER = "Times 20"
    INPUT = "Times 14"

    # Name list
    NAMES = nameList()
    NAMES.sort()

    # Yes or No
    YN = ['Yes', 'No']

    # WTD PW
    WTDPW = ['WTD', 'PW']

    # Root
    root = tk.Tk()

    # Sizing
    canvas = tk.Canvas(root, height = HEIGHT, width = WIDTH)
    canvas.pack()

    # Initial frame
    frame = tk.Frame(root, bg='#BA2809', bd=5)
    frame.place(relx=0.5, rely=0.025, relwidth=0.95, relheight=0.95, anchor='n')

    try:
        # Computers
        comp_frame = tk.Frame(frame, bg='#EBEBEB')
        comp_frame.place(relx=0.01, rely=0.025, relwidth=0.175, relheight=0.7)

        comp_header = tk.Label(comp_frame, bg='#EBEBEB', font=HEADER, text='Computers', fg='black')
        comp_header.place(relx=0, rely=0, relwidth=1, relheight=0.1)

        comp_label_name = tk.Label(comp_frame, bg='#EBEBEB', text='Name:', font=INPUT, fg='black')
        comp_label_name.place(relx=0, rely=0.1, relwidth=0.45, relheight=0.1)

        comp_name = tk.StringVar(comp_frame)
        comp_name.set('Choose')

        comp_name_drop = tk.OptionMenu(comp_frame, comp_name, *NAMES)
        comp_name_drop.place(relx=0.45, rely=0.15, relwidth=0.53, relheight=0.08, anchor='w')

        comp_label_price = tk.Label(comp_frame, bg='#EBEBEB', text='Price:', font=INPUT, fg='black')
        comp_label_price.place(relx=0, rely=0.2, relwidth=0.45, relheight=0.1)

        comp_price = tk.Entry(comp_frame)
        comp_price.place(relx=0.45, rely=0.25, relwidth=0.53, relheight=0.08, anchor='w')

        comp_label_sub = tk.Label(comp_frame, bg='#EBEBEB', text='Subtotal:', font=INPUT, fg='black')
        comp_label_sub.place(relx=0, rely=0.3, relwidth=0.45, relheight=0.1)

        comp_sub = tk.Entry(comp_frame)
        comp_sub.place(relx=0.45, rely=0.35, relwidth=0.53, relheight=0.08, anchor='w')

        comp_label_esp = tk.Label(comp_frame, bg='#EBEBEB', text='ESP:', font=INPUT, fg='black')
        comp_label_esp.place(relx=0, rely=0.4, relwidth=0.45, relheight=0.1)

        comp_esp = tk.StringVar(comp_frame)
        comp_esp.set('Choose')

        comp_esp_drop = tk.OptionMenu(comp_frame, comp_esp, *YN)
        comp_esp_drop.place(relx=0.45, rely=0.45, relwidth=0.53, relheight=0.08, anchor='w')

        comp_label_vs = tk.Label(comp_frame, bg='#EBEBEB', text='VS Plus:', font=INPUT, fg='black')
        comp_label_vs.place(relx=0, rely=0.5, relwidth=0.45, relheight=0.1)

        comp_vs = tk.StringVar(comp_frame)
        comp_vs.set('Choose')

        comp_vs_drop = tk.OptionMenu(comp_frame, comp_vs, *YN)
        comp_vs_drop.place(relx=0.45, rely=0.55, relwidth=0.53, relheight=0.08, anchor='w')

        comp_response = tk.Label(comp_frame, bg='#F5F5F5', fg='black')
        comp_response.place(relx=0.5, rely=0.85, relwidth=0.9, relheight=0.13, anchor='n')

        comp_submit = tk.Button(comp_frame, text='Submit', command=lambda: compSubmit())
        comp_submit.place(relx=0.5, rely=0.625, relwidth=.7, relheight=0.06, anchor='n')

        comp_clear = tk.Button(comp_frame, text='Clear', command=lambda: compClear())
        comp_clear.place(relx=0.5, rely=0.7, relwidth=.7, relheight=0.06, anchor='n')

        comp_ret = tk.Button(comp_frame, text='Submit Return', command=lambda: computerRetIn())
        comp_ret.place(relx=0.5, rely=0.775, relwidth=.7, relheight=0.06, anchor='n')

        #printers
        print_frame = tk.Frame(frame, bg='#EBEBEB')
        print_frame.place(relx=0.21, rely=0.025, relwidth=0.175, relheight=0.7)

        print_header = tk.Label(print_frame, bg='#EBEBEB', font=HEADER, text='Printers', fg='black')
        print_header.place(relx=0, rely=0, relwidth=1, relheight=0.1)

        print_label_name = tk.Label(print_frame, bg='#EBEBEB', text='Name:', font=INPUT, fg='black')
        print_label_name.place(relx=0, rely=0.1, relwidth=0.45, relheight=0.1)

        print_name = tk.StringVar(print_frame)
        print_name.set('Choose')

        print_name_drop = tk.OptionMenu(print_frame, print_name, *NAMES)
        print_name_drop.place(relx=0.45, rely=0.15, relwidth=0.53, relheight=0.08, anchor='w')

        print_label_price = tk.Label(print_frame, bg='#EBEBEB', text='Price:', font=INPUT, fg='black')
        print_label_price.place(relx=0, rely=0.2, relwidth=0.45, relheight=0.1)

        print_price = tk.Entry(print_frame)
        print_price.place(relx=0.45, rely=0.25, relwidth=0.53, relheight=0.08, anchor='w')

        print_label_sub = tk.Label(print_frame, bg='#EBEBEB', text='Subtotal:', font=INPUT, fg='black')
        print_label_sub.place(relx=0, rely=0.3, relwidth=0.45, relheight=0.1)

        print_sub = tk.Entry(print_frame)
        print_sub.place(relx=0.45, rely=0.35, relwidth=0.53, relheight=0.08, anchor='w')

        print_label_esp = tk.Label(print_frame, bg='#EBEBEB', text='ESP:', font=INPUT, fg='black')
        print_label_esp.place(relx=0, rely=0.4, relwidth=0.45, relheight=0.1)

        print_esp = tk.StringVar(print_frame)
        print_esp.set('Choose')

        print_esp_drop = tk.OptionMenu(print_frame, print_esp, *YN)
        print_esp_drop.place(relx=0.45, rely=0.45, relwidth=0.53, relheight=0.08, anchor='w')

        print_label_ink = tk.Label(print_frame, bg='#EBEBEB', text='Ink/Toner:', font=INPUT, fg='black')
        print_label_ink.place(relx=0, rely=0.5, relwidth=0.45, relheight=0.1)

        print_ink = tk.StringVar(print_frame)
        print_ink.set('Choose')

        print_ink_drop = tk.OptionMenu(print_frame, print_ink, *YN)
        print_ink_drop.place(relx=0.45, rely=0.55, relwidth=0.53, relheight=0.08, anchor='w')

        print_response = tk.Label(print_frame, bg='#F5F5F5', fg='black')
        print_response.place(relx=0.5, rely=0.85, relwidth=0.9, relheight=0.13, anchor='n')

        print_submit = tk.Button(print_frame, text='Submit', command=lambda: printSubmit())
        print_submit.place(relx=0.5, rely=0.625, relwidth=.7, relheight=0.06, anchor='n')

        print_clear = tk.Button(print_frame, text='Clear', command=lambda: printClear())
        print_clear.place(relx=0.5, rely=0.7, relwidth=.7, relheight=0.06, anchor='n')

        print_ret = tk.Button(print_frame, text='Submit Return', command=lambda: printerRetIn())
        print_ret.place(relx=0.5, rely=0.775, relwidth=.7, relheight=0.06, anchor='n')

        #chairs
        chair_frame = tk.Frame(frame, bg='#EBEBEB')
        chair_frame.place(relx=0.41, rely=0.025, relwidth=0.175, relheight=0.7)

        chair_header = tk.Label(chair_frame, bg='#EBEBEB', font=HEADER, text='Chairs', fg='black')
        chair_header.place(relx=0, rely=0, relwidth=1, relheight=0.1)

        chair_label_name = tk.Label(chair_frame, bg='#EBEBEB', text='Name:', font=INPUT, fg='black')
        chair_label_name.place(relx=0, rely=0.1, relwidth=0.45, relheight=0.1)

        chair_name = tk.StringVar(chair_frame)
        chair_name.set('Choose')

        chair_name_drop = tk.OptionMenu(chair_frame, chair_name, *NAMES)
        chair_name_drop.place(relx=0.45, rely=0.15, relwidth=0.53, relheight=0.08, anchor='w')

        chair_label_price = tk.Label(chair_frame, bg='#EBEBEB', text='Price:', font=INPUT, fg='black')
        chair_label_price.place(relx=0, rely=0.2, relwidth=0.45, relheight=0.1)

        chair_price = tk.Entry(chair_frame)
        chair_price.place(relx=0.45, rely=0.25, relwidth=0.53, relheight=0.08, anchor='w')

        chair_label_sub = tk.Label(chair_frame, bg='#EBEBEB', text='Subtotal:', font=INPUT, fg='black')
        chair_label_sub.place(relx=0, rely=0.3, relwidth=0.45, relheight=0.1)

        chair_sub = tk.Entry(chair_frame)
        chair_sub.place(relx=0.45, rely=0.35, relwidth=0.53, relheight=0.08, anchor='w')

        chair_label_esp = tk.Label(chair_frame, bg='#EBEBEB', text='ESP:', font=INPUT, fg='black')
        chair_label_esp.place(relx=0, rely=0.4, relwidth=0.45, relheight=0.1)

        chair_esp = tk.StringVar(chair_frame)
        chair_esp.set('Choose')

        chair_esp_drop = tk.OptionMenu(chair_frame, chair_esp, *YN)
        chair_esp_drop.place(relx=0.45, rely=0.45, relwidth=0.53, relheight=0.08, anchor='w')

        chair_label_mat = tk.Label(chair_frame, bg='#EBEBEB', text='Chair Mat:', font=INPUT, fg='black')
        chair_label_mat.place(relx=0, rely=0.5, relwidth=0.45, relheight=0.1)

        chair_mat = tk.StringVar(chair_frame)
        chair_mat.set('Choose')

        chair_mat_drop = tk.OptionMenu(chair_frame, chair_mat, *YN)
        chair_mat_drop.place(relx=0.45, rely=0.55, relwidth=0.53, relheight=0.08, anchor='w')

        chair_response = tk.Label(chair_frame, bg='#F5F5F5', fg='black')
        chair_response.place(relx=0.5, rely=0.85, relwidth=0.9, relheight=0.13, anchor='n')

        chair_submit = tk.Button(chair_frame, text='Submit', command=lambda: chairSubmit())
        chair_submit.place(relx=0.5, rely=0.625, relwidth=.7, relheight=0.06, anchor='n')

        chair_clear = tk.Button(chair_frame, text='Clear', command=lambda: chairClear())
        chair_clear.place(relx=0.5, rely=0.7, relwidth=.7, relheight=0.06, anchor='n')

        chair_ret = tk.Button(chair_frame, text='Submit Return', command=lambda: chairRetIn())
        chair_ret.place(relx=0.5, rely=0.775, relwidth=.7, relheight=0.06, anchor='n')

        #morning reports
        #unit height = 0.0725

        morn_frame = tk.Frame(frame, bg='#EBEBEB')
        morn_frame.place(relx=0.61, rely=0.025, relwidth=0.175, relheight=0.965)

        morn_header = tk.Label(morn_frame, bg='#EBEBEB', font=HEADER, text='Morning Reports', fg='black')
        morn_header.place(relx=0, rely=0, relwidth=1, relheight=0.0725)

        morn_label_conv = tk.Label(morn_frame, bg='#EBEBEB', text='Conversion:', font=INPUT, fg='black')
        morn_label_conv.place(relx=0, rely=0.0725, relwidth=0.45, relheight=0.0725)

        morn_conv = tk.Entry(morn_frame)
        morn_conv.place(relx=0.45, rely=0.10875, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_dpt = tk.Label(morn_frame, bg='#EBEBEB', text='DPT:', font=INPUT, fg='black')
        morn_label_dpt.place(relx=0, rely=0.145, relwidth=0.45, relheight=0.0725)

        morn_dpt = tk.Entry(morn_frame)
        morn_dpt.place(relx=0.45, rely=0.18125, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_prod = tk.Label(morn_frame, bg='#EBEBEB', text='Productive:', font=INPUT, fg='black')
        morn_label_prod.place(relx=0, rely=0.2175, relwidth=0.45, relheight=0.0725)

        morn_prod = tk.Entry(morn_frame)
        morn_prod.place(relx=0.45, rely=0.25375, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_comp = tk.Label(morn_frame, bg='#EBEBEB', text='Comp:', font=INPUT, fg='black')
        morn_label_comp.place(relx=0, rely=0.29, relwidth=0.45, relheight=0.0725)

        morn_comp = tk.Entry(morn_frame)
        morn_comp.place(relx=0.45, rely=0.32625, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_tran = tk.Label(morn_frame, bg='#EBEBEB', text='# of Trans:', font=INPUT, fg='black')
        morn_label_tran.place(relx=0, rely=0.3625, relwidth=0.45, relheight=0.0725)

        morn_tran = tk.Entry(morn_frame)
        morn_tran.place(relx=0.45, rely=0.39875, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_espu = tk.Label(morn_frame, bg='#EBEBEB', text='ESP Units:', font=INPUT, fg='black')
        morn_label_espu.place(relx=0, rely=0.435, relwidth=0.45, relheight=0.0725)

        morn_espu = tk.Entry(morn_frame)
        morn_espu.place(relx=0.45, rely=0.47125, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_esp = tk.Label(morn_frame, bg='#EBEBEB', text='ESP %:', font=INPUT, fg='black')
        morn_label_esp.place(relx=0, rely=0.5075, relwidth=0.45, relheight=0.0725)

        morn_esp = tk.Entry(morn_frame)
        morn_esp.place(relx=0.45, rely=0.54375, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_espd = tk.Label(morn_frame, bg='#EBEBEB', text='ESP $:', font=INPUT, fg='black')
        morn_label_espd.place(relx=0, rely=0.58, relwidth=0.45, relheight=0.0725)

        morn_espd = tk.Entry(morn_frame)
        morn_espd.place(relx=0.45, rely=0.61625, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_etd = tk.Label(morn_frame, bg='#EBEBEB', text='ET $:', font=INPUT, fg='black')
        morn_label_etd.place(relx=0, rely=0.6525, relwidth=0.45, relheight=0.0725)

        morn_etd = tk.Entry(morn_frame)
        morn_etd.place(relx=0.45, rely=0.68875, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_ma = tk.Label(morn_frame, bg='#EBEBEB', text='More Acc:', font=INPUT, fg='black')
        morn_label_ma.place(relx=0, rely=0.725, relwidth=0.45, relheight=0.0725)

        morn_ma = tk.Entry(morn_frame)
        morn_ma.place(relx=0.45, rely=0.76125, relwidth=0.53, relheight=0.055, anchor='w')

        morn_label_nps = tk.Label(morn_frame, bg='#EBEBEB', text='NPS:', font=INPUT, fg='black')
        morn_label_nps.place(relx=0, rely=0.7975, relwidth=0.45, relheight=0.0725)

        morn_nps = tk.Entry(morn_frame)
        morn_nps.place(relx=0.45, rely=0.83375, relwidth=0.53, relheight=0.055, anchor='w')

        morn_response = tk.Label(morn_frame, bg='#F5F5F5', fg='black')
        morn_response.place(relx=0.5, rely=0.935, relwidth=0.9, relheight=0.05, anchor='n')

        morn_submit = tk.Button(morn_frame, text='Submit', command=lambda: mornSubmit())
        morn_submit.place(relx=0.52, rely=0.9, relwidth=.4, relheight=0.048, anchor='w')

        morn_clear = tk.Button(morn_frame, text='Clear', command=lambda: mornClear())
        morn_clear.place(relx=0.48, rely=0.9, relwidth=.4, relheight=0.048, anchor='e')

        #summary
        sum_frame = tk.Frame(frame, bg='#EBEBEB')
        sum_frame.place(relx=0.81, rely=0.025, relwidth=0.175, relheight=0.965)

        sum_header = tk.Label(sum_frame, bg='#EBEBEB', font=HEADER, text='Summary', fg='black')
        sum_header.place(relx=0, rely=0, relwidth=0.65, relheight=0.0725)

        wtd_pw = tk.StringVar(sum_frame)
        wtd_pw.set(WTDPW[0])

        wtd_pw_drop = tk.OptionMenu(sum_frame, wtd_pw, *WTDPW)
        wtd_pw_drop.place(relx=0.65, rely=0.015, relwidth=0.3, relheight=0.05, anchor='nw')

        sum_label_conv = tk.Label(sum_frame, bg='#EBEBEB', text='Conversion:', font=INPUT, fg='black')
        sum_label_conv.place(relx=0, rely=0.0725, relwidth=0.45, relheight=0.0725)

        sum_conv = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_conv.place(relx=0.45, rely=0.10875, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_dpt = tk.Label(sum_frame, bg='#EBEBEB', text='DPT:', font=INPUT, fg='black')
        sum_label_dpt.place(relx=0, rely=0.145, relwidth=0.45, relheight=0.0725)

        sum_dpt = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_dpt.place(relx=0.45, rely=0.18125, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_prod = tk.Label(sum_frame, bg='#EBEBEB', text='Productive:', font=INPUT, fg='black')
        sum_label_prod.place(relx=0, rely=0.2175, relwidth=0.45, relheight=0.0725)

        sum_prod = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_prod.place(relx=0.45, rely=0.25375, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_comp = tk.Label(sum_frame, bg='#EBEBEB', text='Comp:', font=INPUT, fg='black')
        sum_label_comp.place(relx=0, rely=0.29, relwidth=0.45, relheight=0.0725)

        sum_comp = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_comp.place(relx=0.45, rely=0.32625, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_tran = tk.Label(sum_frame, bg='#EBEBEB', text='# of Trans:', font=INPUT, fg='black')
        sum_label_tran.place(relx=0, rely=0.3625, relwidth=0.45, relheight=0.0725)

        sum_tran = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_tran.place(relx=0.45, rely=0.39875, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_espu = tk.Label(sum_frame, bg='#EBEBEB', text='ESP Units:', font=INPUT, fg='black')
        sum_label_espu.place(relx=0, rely=0.435, relwidth=0.45, relheight=0.0725)

        sum_espu = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_espu.place(relx=0.45, rely=0.47125, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_espd = tk.Label(sum_frame, bg='#EBEBEB', text='ESP $:', font=INPUT, fg='black')
        sum_label_espd.place(relx=0, rely=0.5075, relwidth=0.45, relheight=0.0725)

        sum_espd = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_espd.place(relx=0.45, rely=0.54375, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_etd = tk.Label(sum_frame, bg='#EBEBEB', text='ET $:', font=INPUT, fg='black')
        sum_label_etd.place(relx=0, rely=0.58, relwidth=0.45, relheight=0.0725)

        sum_etd = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_etd.place(relx=0.45, rely=0.61625, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_ma = tk.Label(sum_frame, bg='#EBEBEB', text='More Acc:', font=INPUT, fg='black')
        sum_label_ma.place(relx=0, rely=0.6525, relwidth=0.45, relheight=0.0725)

        sum_ma = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_ma.place(relx=0.45, rely=0.68875, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_pc = tk.Label(sum_frame, bg='#EBEBEB', text='PC Basket:', font=INPUT, fg='black')
        sum_label_pc.place(relx=0, rely=0.725, relwidth=0.45, relheight=0.0725)

        sum_pc = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_pc.place(relx=0.45, rely=0.76125, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_print = tk.Label(sum_frame, bg='#EBEBEB', text='Print Basket:', font=INPUT, fg='black')
        sum_label_print.place(relx=0, rely=0.7975, relwidth=0.45, relheight=0.0725)

        sum_print = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_print.place(relx=0.45, rely=0.83375, relwidth=0.53, relheight=0.055, anchor='w')

        sum_label_chair = tk.Label(sum_frame, bg='#EBEBEB', text='Chair Basket:', font=INPUT, fg='black')
        sum_label_chair.place(relx=0, rely=0.87, relwidth=0.45, relheight=0.0725)

        sum_chair = tk.Label(sum_frame, bg='#F5F5F5', fg='black')
        sum_chair.place(relx=0.45, rely=0.90625, relwidth=0.53, relheight=0.055, anchor='w')

        sum_update = tk.Button(sum_frame, text='Update', command=lambda: sumUpdate())
        sum_update.place(relx=0.525, rely=0.9425, relwidth=.4, relheight=0.048, anchor='nw')

        sum_daily = tk.Button(sum_frame, text='Daily Summary', command=lambda: printDaily())
        sum_daily.place(relx=0.075, rely=0.9425, relwidth=.4, relheight=0.048, anchor='nw')

        #associate lookup
        look_frame = tk.Frame(frame, bg='#EBEBEB')
        look_frame.place(relx=0.01, rely=0.75, relwidth=0.575, relheight=0.24)

        look_header = tk.Label(look_frame, bg='#EBEBEB', font=HEADER, text='Lookup', justify='left', fg='black')
        look_header.place(relx=0, rely=0, relwidth=.2, relheight=0.3)

        look_label_name = tk.Label(look_frame, bg='#EBEBEB', text='Name:', font=INPUT, fg='black')
        look_label_name.place(relx=0.2, rely=0, relwidth=0.11, relheight=0.3)

        look_name = tk.StringVar(look_frame)
        look_name.set('Choose')

        look_name_drop = tk.OptionMenu(look_frame, look_name, *NAMES)
        look_name_drop.place(relx=0.3, rely=0.15, relwidth=0.125, relheight=0.25, anchor='w')

        look_label_start_date = tk.Label(look_frame, bg='#EBEBEB', text='Start Date:', font=INPUT, fg='black')
        look_label_start_date.place(relx=0.43, rely=0.15, relwidth=0.125, relheight=0.25, anchor='w')

        look_start_date = tkc.DateEntry(look_frame, width=12, bg='#FF2222', fg='black', borderwidth=2)
        look_start_date.place(relx=0.55, rely=0.15, relwidth=0.1, relheight=0.25, anchor='w')

        look_label_end_date = tk.Label(look_frame, bg='#EBEBEB', text='End Date:', font=INPUT, fg='black')
        look_label_end_date.place(relx=0.66, rely=0.15, relwidth=0.125, relheight=0.25, anchor='w')

        look_end_date = tkc.DateEntry(look_frame, width=12, bg='#FF2222', fg='black', borderwidth=2)
        look_end_date.place(relx=0.78, rely=0.15, relwidth=0.1, relheight=0.25, anchor='w')

        look_label_comp_basket = tk.Label(look_frame, bg='#EBEBEB', text='PC Basket:', font=INPUT, fg='black')
        look_label_comp_basket.place(relx=0, rely=0.35, relheight=0.2, relwidth=0.14)

        look_comp_basket = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_comp_basket.place(relx=0.14, rely=0.35, relheight=0.18, relwidth=0.16)

        look_label_comp_esp = tk.Label(look_frame, bg='#EBEBEB', text='ESP Att:', font=INPUT, fg='black')
        look_label_comp_esp.place(relx=0, rely=0.55, relheight=0.2, relwidth=0.14)

        look_comp_esp = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_comp_esp.place(relx=0.14, rely=0.55, relheight=0.18, relwidth=0.16)

        look_label_comp_vs = tk.Label(look_frame, bg='#EBEBEB', text='VS Att:', font=INPUT, fg='black')
        look_label_comp_vs.place(relx=0, rely=0.75, relheight=0.2, relwidth=0.14)

        look_comp_vs = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_comp_vs.place(relx=0.14, rely=0.75, relheight=0.18, relwidth=0.16)

        look_label_print_basket = tk.Label(look_frame, bg='#EBEBEB', text='Printer Basket:', font=INPUT, fg='black')
        look_label_print_basket.place(relx=0.31, rely=0.35, relheight=0.2, relwidth=0.18)

        look_print_basket = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_print_basket.place(relx=0.49, rely=0.35, relheight=0.18, relwidth=0.16)

        look_label_print_esp = tk.Label(look_frame, bg='#EBEBEB', text='ESP Att:', font=INPUT, fg='black')
        look_label_print_esp.place(relx=0.31, rely=0.55, relheight=0.2, relwidth=0.18)

        look_print_esp = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_print_esp.place(relx=0.49, rely=0.55, relheight=0.18, relwidth=0.16)

        look_label_print_vs = tk.Label(look_frame, bg='#EBEBEB', text='Ink/Toner Att:', font=INPUT, fg='black')
        look_label_print_vs.place(relx=0.31, rely=0.75, relheight=0.2, relwidth=0.18)

        look_print_ink = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_print_ink.place(relx=0.49, rely=0.75, relheight=0.18, relwidth=0.16)

        look_label_chair_basket = tk.Label(look_frame, bg='#EBEBEB', text='Chair Basket:', font=INPUT, fg='black')
        look_label_chair_basket.place(relx=0.65, rely=0.35, relheight=0.2, relwidth=0.18)

        look_chair_basket = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_chair_basket.place(relx=0.82, rely=0.35, relheight=0.18, relwidth=0.16)

        look_label_chair_esp = tk.Label(look_frame, bg='#EBEBEB', text='ESP Att:', font=INPUT, fg='black')
        look_label_chair_esp.place(relx=0.65, rely=0.55, relheight=0.2, relwidth=0.18)

        look_chair_esp = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_chair_esp.place(relx=0.82, rely=0.55, relheight=0.18, relwidth=0.16)

        look_label_chair_vs = tk.Label(look_frame, bg='#EBEBEB', text='Chair Mat Att:', font=INPUT, fg='black')
        look_label_chair_vs.place(relx=0.65, rely=0.75, relheight=0.2, relwidth=0.18)

        look_chair_mat = tk.Label(look_frame, bg='#F5F5F5', fg='black')
        look_chair_mat.place(relx=0.82, rely=0.75, relheight=0.18, relwidth=0.16)

        look_search = tk.Button(look_frame, text='Search', command=lambda: lookSearch(lookName()))
        look_search.place(relx=0.9, rely=0.075, relwidth=0.08, relheight=0.125, anchor='w')

        look_print = tk.Button(look_frame, text='Print All', command=lambda: printAll())
        look_print.place(relx=0.9, rely=0.225, relwidth=0.08, relheight=0.125, anchor='w')

    except:
        name_label = tk.Label(frame, text='Please enter names into associateNames.txt', font='times 50')
        name_label.place(relx=0.5, rely=0, anchor='n', relheight=1, relwidth=1)

    #update summary upon opening
    sumUpdate()

    input('')

if __name__ == '__main__':
    main()
