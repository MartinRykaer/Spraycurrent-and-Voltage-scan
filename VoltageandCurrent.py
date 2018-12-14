# -*- coding: utf-8 -*-
"""
Created on Mon Oct  8 09:41:48 2018

@author: mxg635
"""

import os
import tkinter as tk
from tkinter import filedialog
import datetime as dt
import comtypes, comtypes.client
from ctypes import *
from comtypes.automation import *
import time



def Statuslog(fp): # View/Report/Status Log - retentiontime might be slightly off.
    """
    Made it a function so it can be expanded easier. 
    """
        
    
    pdStatusLogRT = c_double()
    pvarLabels = comtypes.automation.VARIANT()
    pvarValues = comtypes.automation.VARIANT()
    pnArraySize = c_long()
    vclist = []

        
        
    xr = comtypes.client.CreateObject('MSFileReader.XRawfile')
    try:
        xr.open(fp)
    except(OSError):
        return('OSError')
    res = xr.SetCurrentController(0,1)#Needed for some reason.
    ns = c_long()
        
    xr.GetNumSpectra(ns)
        
    #print(ns.value)
        
        
    for i in range(1,ns.value):
        pdStatusLogRT = c_double()
        pvarLabels = comtypes.automation.VARIANT()
        pvarValues = comtypes.automation.VARIANT()
        pnArraySize = c_long()
    
        xr.GetStatusLogForScanNum(c_long(i), byref(pdStatusLogRT), pvarLabels, pvarValues, byref(pnArraySize) )
        print(i)#just to see how far it's getting
        vclist.append([i,pvarValues.value[4], pvarValues.value[5]])
        
        del(pdStatusLogRT, pvarLabels, pvarValues, pnArraySize) #Apparently these have to be deleted and reinitiated every loop iteration...
            
        #for i in range(len(pvarLabels.value)):
        #    savelist.append([i, pvarLabels.value[i]])#, pvarValues.value[i]])
    #print(vclist)
    xr.close()
    return(vclist)


root = tk.Tk()
root.withdraw()

rawfilename = filedialog.askopenfilename(title = "Pick the raw file to parse")
dir_path = filedialog.askdirectory(title = "Pick file directory for savefile")

printlist = Statuslog(rawfilename)

outfile = open(dir_path+'voltagecurrentout2.txt','w+')
for entry in printlist:
    outfile.write(str(entry[0])+'\t'+str(entry[1])+'\t'+str(entry[2])+'\n')
outfile.close()



"""
Entries in statuslog:

[0, '====  Overall Status:  ====:']
[1, 'Status:']
[2, 'Performance:']
[3, '======  Ion Source:  ======:']
[4, 'Spray Voltage (V)']
[5, 'Spray Current (µA)']
[6, 'Spray Current std. dev. (µA)']
[7, 'Capillary Temperature (°C)']
[8, 'Sheath gas flow rate']
[9, 'Aux gas flow rate']
[10, 'Sweep gas flow rate']
[11, 'Aux. Temperature (°C)']
[12, '======  Ion Optics:  ======:']
[13, 'Capillary Voltage (V)']
[14, 'Bent Flatapole DC (V)']
[15, 'Inj Flatapole A DC (V)']
[16, 'Inj Flatapole B DC (V)']
[17, 'Trans Multipole DC (V)']
[18, 'HCD Multipole DC (V)']
[19, 'Inj. Flat. RF Amp (V)']
[20, 'Inj. Flat. RF Freq (kHz)']
[21, 'Bent Flat. RF Amp (V)']
[22, 'Bent Flat. RF Freq (kHz)']
[23, 'RF2 and RF3 Amp (V)']
[24, 'RF2 and RF3 Freq (kHz)']
[25, 'Inter Flatapole DC (V)']
[26, 'Quad Exit DC (V)']
[27, 'C-Trap Entrance Lens DC (V)']
[28, 'C-Trap RF Amp (V)']
[29, 'C-Trap RF Freq (kHz)']
[30, 'C-Trap RF Curr (A)']
[31, 'C-Trap Exit Lens DC (V)']
[32, 'HCD Exit Lens DC (V)']
[33, '======  Vacuum:  ======:']
[34, 'Fore Vacuum Sensor (mbar)']
[35, 'High Vacuum Sensor (mbar)']
[36, 'UHV Sensor (mbar)']
[37, 'Source TMP Speed']
[38, 'UHV TMP Speed']
[39, '=====  Temperatures:  =====:']
[40, 'Analyzer Temperature (°C)']
[41, 'Ambient Temperature (°C)']
[42, 'Ambient Humidity (%)']
[43, 'Source TMP Motor Temperature (°C)']
[44, 'Source TMP Bottom Temperature (°C)']
[45, 'UHV TMP Motor Temperature (°C)']
[46, 'IOS Heatsink Temp. (°C)']
[47, 'HVPS Peltier Temp. (°C)']
[48, 'Quad. Det. Temp. (°C)']
[49, '====  Diagnostic Data:  ====:']
[50, 'Performance ld']
[51, 'Performance me']
[52, 'Performance cy:']
[53, 'CTCD mV']
"""