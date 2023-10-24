# -*- coding: utf-8 -*-
"""
Created on Thu Oct 12 11:30:41 2023

@author: OABI
"""

import os, time, re, sys, inspect, psutil, string, shutil
from datetime import datetime
import win32com.client
import pythoncom
import copy # to create independent copies from objects
import threading # to create background processes
import gc
import pywintypes

# DEFINED CLASSES
from Classes.DS_BES_AI import DS_BES_AI
from Classes.DS_BES_AO import DS_BES_AO
from Classes.DS_BES_Antenna import DS_BES_Antenna
from Classes.DS_BES_Beacon import DS_BES_Beacon
from Classes.DS_BES_Cable import DS_BES_Cable
from Classes.DS_BES_Compass import DS_BES_Compass
from Classes.DS_BES_CV import DS_BES_CV
from Classes.DS_BES_DI import DS_BES_DI
from Classes.DS_BES_DO import DS_BES_DO
from Classes.DS_BES_Enclosure import DS_BES_Enclosure
from Classes.DS_BES_FD import DS_BES_FD
from Classes.DS_BES_FI import DS_BES_FI
from Classes.DS_BES_Fogdetector import DS_BES_Fogdetector
from Classes.DS_BES_Foghorn import DS_BES_Foghorn
from Classes.DS_BES_GD import DS_BES_GD
from Classes.DS_BES_Handsw import DS_BES_Handsw
from Classes.DS_BES_LG import DS_BES_LG
from Classes.DS_BES_Limitsw import DS_BES_Limitsw
from Classes.DS_BES_LIT import DS_BES_LIT
from Classes.DS_BES_Loadpin import DS_BES_Loadpin
from Classes.DS_BES_Oceanograph import DS_BES_Oceanograph
from Classes.DS_BES_PG import DS_BES_PG
from Classes.DS_BES_PIG import DS_BES_PIG
from Classes.DS_BES_PIT import DS_BES_PIT
from Classes.DS_BES_RO import DS_BES_RO
from Classes.DS_BES_SOL_V import DS_BES_SOL_V
from Classes.DS_BES_Solarpanel import DS_BES_Solarpanel
from Classes.DS_BES_Slipring import DS_BES_Slipring
from Classes.DS_BES_Speaker import DS_BES_Speaker
from Classes.DS_BES_SV import DS_BES_SV
from Classes.DS_BES_TG import DS_BES_TG
from Classes.DS_BES_TT import DS_BES_TT
from Classes.DS_BES_Weatherstation import DS_BES_Weatherstation
from Classes.DS_BES_Windgenerator import DS_BES_Windgenerator
from Classes.DS_BES_Transformer import DS_BES_Transformer
from Classes.Instrument import Instrument
from Classes.InstrumentFunction import InstrumentFunction
from Classes.Tooltip import CreateToolTip
# import ZODB and supporting libraries
from ZEO import ClientStorage
from ZODB import FileStorage, DB
import transaction

# global variables
addr = '10.175.13.199', 8091 # This is the address of the ZEO Server
DS_List = os.listdir('Datasheets')


global Instrumentation

Instrumentation = ["DS_BES_Antenna","DS_BES_Beacon","DS_BES_Compass","DS_BES_CV","DS_BES_FD",
                   "DS_BES_FI","DS_BES_Fogdetector","DS_BES_Foghorn","DS_BES_GD","DS_BES_Handsw",
                   "DS_BES_LG","DS_BES_Limitsw","DS_BES_LIT","DS_BES_Loadpin","DS_BES_Oceanograph",
                   "DS_BES_PG","DS_BES_PIG","DS_BES_PIT","DS_BES_RO","DS_BES_SOL_V","DS_BES_SV",
                   "DS_BES_TG","DS_BES_TT","DS_BES_Weatherstation"]

print("LOOPEXPORT")
# Part 1: Collecting required data to produce loops
InstrumentList = []
CableList = []
ConnectionsDict = {}
LSignalDict = {}
LTerminationDict = {}
LSCRDict = {}
ColorDict = {}
RSCRDict = {}
RTerminationDict = {}
RSignalDict = {}
CoreNumberDict = {}
storage = ClientStorage.ClientStorage(addr)             
db = DB(storage)
connection = db.open()
root = connection.root()
for key in root:
    obj = root[key]
    if obj.__class__.__name__ in Instrumentation:
        #print(obj.__class__.__name__)
        InstrumentList.append(key)
#        print(key)
    #progress.update()
    if isinstance(obj,DS_BES_Cable):
        CableList.append(key)
        ConnectionsDict[key] = []
        LSignalDict[key] = []
        LTerminationDict[key] = []
        LSCRDict[key] = []
        ColorDict[key] = []
        RSCRDict[key] = []
        RTerminationDict[key] = []
        RSignalDict[key] = []
        CoreNumberDict[key] = []
        ConnectionsDict[key].append(getattr(obj,"Connection1"))
        ConnectionsDict[key].append(getattr(obj,"Connection2"))
#        print (ConnectionsDict[key])
#            max no of cores is 50
        for i in range(1,51):
           LSignalDict[key].append(getattr(obj,"LSignal"+str(i))) 
           LTerminationDict[key].append(getattr(obj,"LTermination"+str(i)))
           LSCRDict[key].append(getattr(obj,"LSCR"+str(i)))
           ColorDict[key].append(getattr(obj,"Color"+str(i)))
           RSCRDict[key].append(getattr(obj,"RSCR"+str(i)))
           RTerminationDict[key].append(getattr(obj,"RTermination"+str(i)))
           RSignalDict[key].append(getattr(obj,"RSignal"+str(i))) 
           CoreNumberDict[key].append(getattr(obj,"CoreNumber"+str(i)))
           
    elif not isinstance (obj,DS_BES_Cable) and obj.__class__.__name__ not in Instrumentation:
        continue
    
#print (InstrumentList)
#print (ConnectionsDict)  
connection.close()
storage.close()
#print (InstrumentList)
#    print("IC-SPM-LIT-6622-01")
#    print("==================")
#    print(ConnectionsDict["IC-SPM-LIT-6622-01"])
#    print(LSignalDict["IC-SPM-LIT-6622-01"])
#    print(LTerminationDict["IC-SPM-LIT-6622-01"])
#    print(LSCRDict["IC-SPM-LIT-6622-01"])
#    print(ColorDict["IC-SPM-LIT-6622-01"])
#    print(RSCRDict["IC-SPM-LIT-6622-01"])
#    print(RTerminationDict["IC-SPM-LIT-6622-01"])
#    print(RSignalDict["IC-SPM-LIT-6622-01"])
#    print(CoreNumberDict["IC-SPM-LIT-6622-01"])
#    print("IC-SPM-IJB-6620-01")
#    print("==================")
#    print(ConnectionsDict["IC-SPM-IJB-6620-01"])
#    print(LSignalDict["IC-SPM-IJB-6620-01"])
#    print(LTerminationDict["IC-SPM-IJB-6620-01"])
#    print(LSCRDict["IC-SPM-IJB-6620-01"])
#    print(ColorDict["IC-SPM-IJB-6620-01"])
#    print(RSCRDict["IC-SPM-IJB-6620-01"])
#    print(RTerminationDict["IC-SPM-IJB-6620-01"])
#    print(RSignalDict["IC-SPM-IJB-6620-01"])
#    print(CoreNumberDict["IC-SPM-IJB-6620-01"])    
# Part 3: Compile Loops and Preliminary Check
LoopDict = {}
MissingInstrument = []
for item in InstrumentList:
    LoopDict[item] = []
    for Cable in ConnectionsDict:
        if ConnectionsDict[Cable][0] == item: # Connection1 equals the Instrument
            LoopDict[item].append(ConnectionsDict[Cable][0]) # Add Connection 1
            LoopDict[item].append(Cable)                     # Add the Cable
            LoopDict[item].append(ConnectionsDict[Cable][1]) # Add Connection 2
            # Loop through the LSignal list to find the Instrument Signal
            index = 0
            for LSignal in LSignalDict[Cable]:   
#                    check if instrument tag is included in the signal tag, then append to loop dict.
                if item in LSignal:
                    LoopDict[item].append(LSignal)
                    LoopDict[item].append(LTerminationDict[Cable][index])
                    LoopDict[item].append(CoreNumberDict[Cable][index])
                    LoopDict[item].append(ColorDict[Cable][index])
                    LoopDict[item].append(CoreNumberDict[Cable][index])
                    LoopDict[item].append(RTerminationDict[Cable][index])
                    LoopDict[item].append(RSignalDict[Cable][index])
                index += 1
            LoopDict[item].append('*')
    print (LoopDict, "+ new loop")
    if len(LoopDict[item]) == 0:
        MissingInstrument.append(item)
if len(MissingInstrument) != 0:
#    print(MissingInstrument)
    InstrumentString = ''
    for item in MissingInstrument:
        InstrumentString += ','+item

#    tk.messagebox.showwarning(title=None, message="Some Instruments are not terminated.\n"+InstrumentString)
## 2nd round
#for item in InstrumentList:
#    for Cable in ConnectionsDict:
#        print (LoopDict[item][2])
#        print (ConnectionsDict[Cable][2])
#        if ConnectionsDict[Cable][0] == LoopDict[item][2]:   # Connection1 equals Connection2
#            # Loop through the LSignal list to find the Instrument Signal
#            Flag = True
#            index = 0
#            for LSignal in LSignalDict[Cable]:
#                if item in LSignal:
#                    if item == 'SPM-HS-3501A':
#                        print(LSignal)
#                        print(index)
#                    if Flag:
#                        LoopDict[item].append(ConnectionsDict[Cable][0]) # Add Connection 1
#                        LoopDict[item].append(Cable)                     # Add the Cable
#                        LoopDict[item].append(ConnectionsDict[Cable][1]) # Add Connection 2
#                        Flag = False
#                    LoopDict[item].append(LSignal)
#                    LoopDict[item].append(LTerminationDict[Cable][index])
#                    LoopDict[item].append(CoreNumberDict[Cable][index])
#                    LoopDict[item].append(ColorDict[Cable][index])                        
#                    LoopDict[item].append(CoreNumberDict[Cable][index])
#                    LoopDict[item].append(RTerminationDict[Cable][index])
#                    LoopDict[item].append(RSignalDict[Cable][index])
#                index += 1
#            LoopDict[item].append('+')
#print(LoopDict["SPM-LIT-6622"])
##    # 3rd round
##    for item in InstrumentList:
##        print(LoopDict[item][-1:])
##        if LoopDict[item][-1:] != '*':
##            i = LoopDict[item].index('*')+3
##            print(item,LoopDict[item][i])
##            for Cable in ConnectionsDict:
##                if ConnectionsDict[Cable][0] == LoopDict[item][i]:  # Connection1 equals Connection2
##                    # Loop through the LSignal list to find the Instrument Signal
##                    Flag = True
##                    index = 0
##                    for LSignal in LSignalDict[Cable]:
##                        if item in LSignal:
##                            if item == 'SPM-HS-3501A':
##                                print(LSignal)
##                                print(index)
##                            if Flag:
##                                LoopDict[item].append(ConnectionsDict[Cable][0]) # Add Connection 1
##                                LoopDict[item].append(Cable)                     # Add the Cable
##                                LoopDict[item].append(ConnectionsDict[Cable][1]) # Add Connection 2
##                                Flag = False
##                            LoopDict[item].append(LSignal)
##                            LoopDict[item].append(LTerminationDict[Cable][index])
##                            LoopDict[item].append(CoreNumberDict[Cable][index])
##                            LoopDict[item].append(ColorDict[Cable][index])                        
##                            LoopDict[item].append(CoreNumberDict[Cable][index])
##                            LoopDict[item].append(RTerminationDict[Cable][index])
##                            LoopDict[item].append(RSignalDict[Cable][index])
##                        index += 1
##                LoopDict[item].append('&')
path = os.getcwd()+"\Loops Export.xlsx"
## Remove exisiting collection and create new one
if os.path.exists(path):
    os.remove(path)
time.sleep(2) # To give the os module time to remove the Excel Export File    
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Loop = Xcel.Workbooks.Add() 
Loop.SaveAs(path)
row = 2
for item in InstrumentList:
#    Export.ActiveSheet.Cells(row,1).Value = item
#    row += 1
    for dictionary in LoopDict:
        if item == key:
            Loop.ActiveSheet.Cells(row,1)= LoopDict[key][0]
            column +=1
            for i in LoopDict[key]:
                if i != item:
                   Loop.ActiveSheet.Cells(row,1) 
        
#    LoopString = ','.join(map(str,LoopDict[item]))
#    Export.ActiveSheet.Cells(row,1).Value = LoopString.split('*')[0]
#    row += 1
#    Export.ActiveSheet.Cells(row,1).Value = LoopString.split(',')[1]
#    row += 1
Export.ActiveSheet.Range("A:BM").Columns.AutoFit() 
Export.Close(SaveChanges=True)        
Xcel.Application.Quit()