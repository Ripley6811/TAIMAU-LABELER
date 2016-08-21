# -*- coding: utf-8 -*-
'''
Program for printing to a TSC TTP-243E Plus label printer.
Coding information and examples at:
http://www.tscprinters.com/cms/upload/download_en/DLL_instruction.pdf

Date completed: 2013-11-15

Updated: 2013-11-29
-Changed max number to print function: Exits when blank
-Verification of RT. No.: 10 digits and uniqueness

Updated: 2014-9-20
-Printer name changed slightly. Now searches for printer by keyword in name.
-Program will ask for input again instead of exiting on errors.

Updated: 2014-9-22
-Must avoid doing any kind of unnecessary encoding/decoding.

Updated: 2016-7-8
-Now looks for shared network printers also.

Updated: 2016-8-21
-Now gets data from network server through urllib2 requests.

'''
import urllib2
import json

#from ctypes import cdll
from datetime import date
from time import sleep
from win32print import EnumPrinters, PRINTER_ENUM_CONNECTIONS, PRINTER_ENUM_LOCAL
from TscLib import TscLib


''' FIND PRINTER PORTNAME AND INIT TSC_LIB '''
portkeyword="243" # "TSC TTP-243 Plus"
portname = None
printers = [p[2] for p
            in EnumPrinters(PRINTER_ENUM_CONNECTIONS + PRINTER_ENUM_LOCAL)
            if portkeyword in p[2]]

if len(printers) == 1:
    portname, = printers # Asserts only one printer is in list.
    print "Printer:", portname
elif len(printers) > 1:
    for i, printer in enumerate(printers):
        print '{})  {}'.format(i+1, printer)

    # Get user selected index.
    printerIndex = input("Select row number:")-1
    portname = printers[printerIndex]
    print "Printer:", portname

else:
    print "No printer found with keyword '{}'.".format(portkeyword)

tsclib = TscLib(portname) if portname else None



def TM_label(material, PN, LOT_NO, ASE_NO, QTY, ExpDate, DOM, RT_NO, PO):
    if not tsclib:
        return
    tsclib.openport()
    tsclib.setup()
    tsclib.clearbuffer()

    tab = 35
    tab2 = 154
    tab3 = 408

    noPNadjust = 0
    if not PN:
        noPNadjust = 43

    tsclib.windowsfont(tab,15+noPNadjust, "Material name:", 30)
    tsclib.windowsfont(tab+180,6+noPNadjust, material.encode('big5'), 42, style=2)

    if PN:
        tsclib.windowsfont(tab,58, "P/N:", h=26)
        tsclib.windowsfont(tab2,55, PN, h=30, style=2)
        tsclib.windowsfont(tab3,55, u"料號".encode('big5'), h=30)
        tsclib.barcode(tab,85, PN, d=40)

    tsclib.windowsfont(tab,128, "LOT NO:", h=26)
    tsclib.windowsfont(tab2,125, LOT_NO, h=30, style=2)
    tsclib.windowsfont(tab3,125, u"批號".encode('big5'), h=30)
    tsclib.barcode(tab,155, LOT_NO, d=40)

    if PN:
        #windowsfont(tab,198, "ASE No:".encode('big5'), h=26)
        tsclib.windowsfont(tab,198, "Min Pkg No:".encode('big5'), h=24)
    else:
        tsclib.windowsfont(tab,198, "BOX NO:".encode('big5'), h=26)

#    windowsfont(tab,198, "BOX NO:".encode('big5'), h=26)
    tsclib.windowsfont(tab2,195, ASE_NO, h=30, style=2)
    tsclib.barcode(tab,225, ASE_NO, d=40)

#    if isinstance(QTY, str):
#        QTY = QTY.encode('big5')
    tsclib.windowsfont(tab,268, "Q'TY:", h=26)
    tsclib.windowsfont(tab2,265, QTY, h=30, style=2)
    tsclib.windowsfont(tab3,265, u"容量".encode('big5'), h=30)
    tsclib.barcode(tab,295, QTY, d=40)

    tsclib.windowsfont(tab,338, "Exp Date:", h=26)
    tsclib.windowsfont(tab2,335, ExpDate, h=30, style=2)
    tsclib.windowsfont(tab3,335, u"使用期限".encode('big5'), h=30)
    tsclib.barcode(tab,365, ExpDate, d=30)

    tsclib.windowsfont(tab,398, "DOM:", h=26)
    tsclib.windowsfont(tab2,395, DOM, h=30, style=2)
    tsclib.windowsfont(tab3,395, u"製造日期".encode('big5'), h=30)
    tsclib.barcode(tab,425, DOM, d=30)

    if RT_NO:
        tsclib.windowsfont(tab,458, "RT NO:", h=26)
        tsclib.windowsfont(tab2,455, RT_NO, h=30, style=2)
        tsclib.barcode(tab,485, RT_NO, d=40)
    elif PO:
        tsclib.windowsfont(tab,458, "PO:", h=26)
        tsclib.windowsfont(tab2,455, PO, h=30, style=2)
        tsclib.barcode(tab,485, PO, d=40)

    tsclib.printlabel(1,1)
    tsclib.closeport()
    sleep(0.5)


def TM_DMlabel(material, PN, LOT_NO, ASE_NO, QTY, ExpDate, DOM, RT_NO):
    tsclib.openport()
    tsclib.setup()
    tsclib.clearbuffer()

    tab = 145
    tab2 = 262

    BOX_NO = str(int(ASE_NO[-4:]))
    datastr = ('"'+str(RT_NO) + "|" +
              str(PN) + "|" +
              str(LOT_NO) + "|" +
              str(int(QTY)) + "|" +
              str(DOM) + "|" +
              str(BOX_NO)+'"')

    '''The default 'model' does not work. Use 'M2' for enhanced version.
    '''
    #sendcommand('QRCODE 10,10,Q,7,A,0,M1,S0,"M1,S0 THE FIRMWARE HAS BEEN UPDATED"')
    #sendcommand('QRCODE 10,300,Q,7,A,0,M2,S0,"M2,S0 THE FIRMWARE HAS BEEN UPDATED"')
    #sendcommand('QRCODE 300,300,Q,7,A,0,M2,S1,"M2,S1 THE FIRMWARE HAS BEEN UPDATED"')
    tsclib.sendcommand('DMATRIX 190,40,400,400,'+datastr)


    tsclib.windowsfont(tab,268, "BATCH:", h=26)
    tsclib.windowsfont(tab2,265, RT_NO, h=30, style=2)

    tsclib.windowsfont(tab,308, "Part No:", h=26)
    tsclib.windowsfont(tab2,305, PN, h=30, style=2)

    tsclib.windowsfont(tab,348, "Lot No:", h=26)
    tsclib.windowsfont(tab2,345, LOT_NO, h=30, style=2)

    tsclib.windowsfont(tab,388, "Qty:", h=26)
    tsclib.windowsfont(tab2,385, str(int(QTY)), h=30, style=2)

    tsclib.windowsfont(tab,428, "MFG Date:", h=26)
    tsclib.windowsfont(tab2,425, DOM, h=30, style=2)

    tsclib.windowsfont(tab,468, "Box No:", h=26)
    tsclib.windowsfont(tab2,465, BOX_NO, h=30, style=2)


    tsclib.printlabel(1,1)
    tsclib.closeport()
    sleep(0.5)

def TM_QRlabel(material, PN, LOT_NO, ASE_NO, QTY, ExpDate, DOM, RT_NO):
    tsclib.openport()
    tsclib.setup()
    tsclib.clearbuffer()

    '''The default 'model' does not work. Use 'M2' for enhanced version.
    '''
    tsclib.sendcommand('QRCODE 10,300,Q,7,A,0,M2,S0,"M2,S0 THE FIRMWARE HAS BEEN UPDATED"')
    tsclib.sendcommand('QRCODE 300,300,Q,7,A,0,M2,S1,"M2,S1 THE FIRMWARE HAS BEEN UPDATED"')

    tsclib.printlabel(1,1)
    tsclib.closeport()
    sleep(0.5)




def spacedString(*args):
    return bufferString(' ', *args)

def zeroesString(*args):
    return bufferString('0', *args)

def bufferString(bufferChar = '', string = '', length = 0, postBuffer = False):
    strLen = len(u'{}'.format(string))
    if postBuffer:
        return u'{}{}'.format(string, bufferChar * (length-strLen))
    return u'{}{}'.format(bufferChar * (length-strLen), str(string))



'''CMD Interface for working with 桶裝出貨表 data when running this module
independently.
'''
def printapp(noprint=False):

    # Get list of company names and get user selection.
    response = urllib2.urlopen(
        'http://192.168.1.123:8/api/barrelShipment/companies'
    )
    companyNames = json.loads(response.read())

    for i, co in enumerate(companyNames):
        print '{})'.format(i+1),
        print spacedString(co, 6)

    companyIndex = input("Select row number:")-1
    company = companyNames[companyIndex]

    # Get recent shipments and get user selection
    response = urllib2.urlopen(
        'http://192.168.1.123:8/api/barrelShipment/latest/'+company+'/30'
    )
    shipmentArray = json.loads(response.read())

    for i in range(len(shipmentArray))[::-1]:
        sh = shipmentArray[i]
        print '{})'.format(i+1),
        print spacedString('{0}/{1}'.format(sh['shipMonth']+1,sh['shipDate']), 5),
        print spacedString(sh['formID'], 7),
        print spacedString(sh['company'], 7),
        print spacedString(sh['lotID'], 10),
        print zeroesString(sh['start'], 4) + '-' + zeroesString(sh['start']+sh['count'], 4),
        print spacedString('({})'.format(i+1), 5),
        print spacedString(sh['count'], 3),
        print spacedString(sh['product'], 10, True)

    shipmentIndex = input("Select row number:")-1
    doc = shipmentArray[shipmentIndex]


    print 'PRODUCT:', doc[u"product"]
    print 'PN:', doc[u"pn"]
    print doc[u"count"], "labels in this set."
    print ''


    # Get DOM and Exp date
    ExpDate = "dddddddd"
    pdate = doc[u"lotID"][1:7]
    dateDOM = date(2000+int(pdate[:2]), int(pdate[2:4]), int(pdate[4:6]))
    DOM = "{0:04}{1:02}{2:02}".format(dateDOM.year, dateDOM.month, dateDOM.day)
    try:
        inc_year = False
        if int((dateDOM.month-1 + doc[u"shelfLife"]) / 12):
            inc_year = True
        ExpDate = "{0:04}{1:02}{2:02}".format(
                (dateDOM.year + int((dateDOM.month-1 + doc[u"shelfLife"]) / 12)) if inc_year else dateDOM.year,
                int((dateDOM.month-1 + doc[u"shelfLife"]) % 12)+1,
                dateDOM.day)
    except:
        raw_input('Error converting expiration date. Using "dddddddd"\nHIT ENTER TO CONTINUE...')


    # Give opportunity to print again with a while loop.
    while True:
        print "Set max number to print (BLANK will EXIT)."
        nPrint = raw_input("Stop at #")
        try:
            nPrint = int(nPrint)
        except:
            print "Not a valid number... Exiting."
            break

        # Loop through and print each label.
        for i in range(int(doc[u"count"])):
            if i >= nPrint:
                break
            ASE_NO = "{0}{1:04}".format(doc[u"lotID"],doc[u"start"] + i)
            print "---------------------"
            try:
                print doc[u"product"]
            except:
                print "(can't print a Chinese character)"
            print "PN:", doc[u"pn"]
            print "LOT#:", doc[u"lotID"]
            print "ASE#:", ASE_NO
            print "QTY:", doc[u"count"]
            print "EXP:", ExpDate
            print "DOM:", DOM
            print "RT:", doc[u"rtCode"]
            print "PO:", doc[u"orderID"]
            print "---------------------"
            if not noprint:
                if doc['barcode']:
                    TM_label(doc[u"product"], doc[u"pn"], doc[u"lotID"], ASE_NO,
                             doc[u"count"], ExpDate, DOM, doc[u"rtCode"], doc[u"orderID"])
                if doc['datamatrix']:
                    TM_DMlabel(doc[u"product"], doc[u"pn"], doc[u"lotID"], ASE_NO,
                             doc[u"count"], ExpDate, DOM, doc[u"rtCode"])
        if noprint:
            raw_input("Finished. Press enter to close")
    raw_input("Hit enter to close")



if __name__ == '__main__':
    printapp()
