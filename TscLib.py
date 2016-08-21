#!/usr/bin/env python
# -*- coding: utf-8 -*-
"""
summary

description

:REQUIRES:

:TODO:

:AUTHOR: Ripley6811
:ORGANIZATION: None
:CONTACT: python@boun.cr
:SINCE: Sun Aug 21 13:49:01 2016
:VERSION: 0.1
"""
#===============================================================================
# PROGRAM METADATA
#===============================================================================
__author__ = 'Ripley6811'
__contact__ = 'python@boun.cr'
__copyright__ = ''
__license__ = ''
__date__ = 'Sun Aug 21 13:49:01 2016'
__version__ = '0.1'


from ctypes import cdll


class TscLib:
    def __init__(self, portname):
        self.portname = portname

        #XXX: Ensure the TSCLIB.dll is located in windows/system32 folder.
        try:
            self.tsc = cdll.LoadLibrary("TSCLIB.DLL")
        except WindowsError as e:
            print "Failed to load TSCLIB.DLL:", e
            self.tsc = None

    #XXX: Portname must match the name of the printer on the system.
    def openport(self):
        try:
            self.tsc.openport(self.portname)
        except ValueError:
            #XXX: ValueError always thrown but functions properly. Ignore it.
            pass
        print 'PORT "{}" OPENED'.format(self.portname)

    def setup(self, w="70",h="70",c="2",d="2",e="0",f="3",g="0"):
        # w = label width mm
        # h = label height mm
        # c = print speed
        # d = print density
        # e = sensor type
        # f = vertical gap height in mm
        # g = horizontal gap shift
        try:
            self.tsc.setup(str(w),str(h),str(c),str(d),str(e),str(f),str(g))
        except ValueError:
            pass

    # clear all previously entered barcode and text
    def clearbuffer(self):
        self.tsc.clearbuffer()

    # Must pass strings for all values
    def barcode(self, x,y,text,d="40",c="128",e="0",f="0",g="2",h="4"):
        try:
            # x, y starting point
            # d = height
            self.tsc.barcode(str(x),str(y),str(c),str(d),str(e),str(f),str(g),str(h),str(text))
        except ValueError:
            pass

    def windowsfont(self, x,y,text,h=40,rotation=0,style=0,line=0,font=u"Arial"):
        # x,y starting point
        # h = font height
        # rotation = counter clockwise rotation degrees
        # style: 0 = Normal, 1 = Italic, 2 = Bold, 3 = Italic Bold
        # line: 0 = no underline, 1 = underline
        # font = Font type face
        # text = Text to print
        try:
            self.tsc.windowsfont(int(x),int(y),int(h),int(rotation),int(style),int(line),str(font),str(text))
        except ValueError:
            pass

    def printlabel(self, a=1, b=1):
        # a = number of label sets
        # b = number of print copies
        try:
            self.tsc.printlabel(str(a),str(b))
        except ValueError:
            pass

    def sendcommand(self, s):
        try:
            self.tsc.sendcommand(s)
        except ValueError:
            pass

    def closeport(self):
        self.tsc.closeport()
