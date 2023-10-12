# -*- coding: utf-8 -*-
"""
Created on Thu Oct 12 11:30:41 2023

@author: OABI
"""

import os, time, re, sys, inspect, psutil, string, shutil
from datetime import datetime
import win32com.client
import pythoncom

from Classes.DS_BES_Enclosure import DS_BES_Enclosure
from Classes.DS_BES_Cable import DS_BES_Cable
from Classes.DS_BES_Slipring import DS_BES_Slipring
from ZEO import ClientStorage
from ZODB import FileStorage, DB

addr = '10.175.13.199', 8090 # This is the address of the ZEO Server
DS_List = os.listdir('Datasheets')


def terminationexport():
    print("TERMINATIONEXPORT")
    global Spares_List, Spares_Terms
    global CableTypeDict
    global addr
    CableDict = {}
    EnclosureDict = {}
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
    progress.start()
    progress.update()      
    try:
        btn_terminationexport.config(state=tk.DISABLED)
        btn_terminationimport.config(state=tk.DISABLED)
#        storage = FileStorage.FileStorage(directory+filename) 
        storage = ClientStorage.ClientStorage(addr)             
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for key in root:
            obj = root[key]
            progress.update()
            """
            obtain the cables, the connections and the number of cores per cable in a list per cable
            list comprehension:
            CableDict[key][0] = connection1
            CableDict[key][1] = Core Configuration - If available else "NO CORE CONFIGURATION YET"
            CableDict[key][2] = connection2
            CableDict[key][3] = Start of LSignal1..50
            CableDict[key][53] = Start of LTermination1..50
            CableDict[key][103 Start of LSCR1..50
            CableDict[key][153 = Start of Color1..50
            CableDict[key][203] Start of RSCR1..50
            CableDict[key][253] = Start of RTermination1..50
            CableDict[key][303] Start of RSignal1..50
            CableDict[key][353 = Start of CoreNumber1..50
            CableDict[key][403] CoreConfiguration e.g. '2Pr. 22 AWG OS'
            """
            if isinstance(obj,DS_BES_Cable):
                CableDict[key] = []
                CableDict[key].append(obj.Connection1)
                """
                The core configuration can be obtained from CableTypeDict.
                Example:
                CableTypeDict[obj.Model] = ['Belden', '973107Z', '2Pr. 22 AWG OS', '300V', 
                '22 AWG', 'LSZH-TS', '15.2mm', 'M20', None, None, None, None, '', 2, 2, 'OS']
                The obj.Model attribute should in this case be 'RFOU4-6/6'
                The last 3 values of this list represent the core configuration.
                By using CableTypeDict[key][-3:] we obtain these 3 values in a list: [2, 2, 'OS']
                Possible formats: [x, y, 'IS'], [x, y, 'G'], [x, y, 'OS'], [x, y, 'PE']
                """                
                if obj.CoreConfiguration != None: # something like: '2Pr. 22 AWG OS' see CableTypes.xlsx 
                    if obj.Model in CableTypeDict:
                        CableDict[key].append(CableTypeDict[obj.Model][-3:])
                    else:
                        CableDict[key].append("NO (CORRECT) MODEL ASSIGNED YET")
                if obj.CoreConfiguration == None:
                    CableDict[key].append("NO CORE CONFIGURATION YET")
                CableDict[key].append(obj.Connection2)
                for i in range(1,51): # maximum number of cores is 50
                    CableDict[key].append(getattr(obj,"LSignal"+str(i)))
                for i in range(1,51):
                    CableDict[key].append(getattr(obj,"LTermination"+str(i)))
                for i in range(1,51):
                    CableDict[key].append(getattr(obj,"LSCR"+str(i)))
                for i in range(1,51):
                    CableDict[key].append(getattr(obj,"Color"+str(i)))
                for i in range(1,51):
                    CableDict[key].append(getattr(obj,"RSCR"+str(i)))
                for i in range(1,51):
                    CableDict[key].append(getattr(obj,"RTermination"+str(i)))
                for i in range(1,51): # maximum number of cores is 50
                    CableDict[key].append(getattr(obj,"RSignal"+str(i)))
                for i in range(1,51): # maximum number of cores is 40
                    CableDict[key].append(getattr(obj,"CoreNumber"+str(i)))                    
                if obj.CoreConfiguration != None:
                    CableDict[key].append(obj.CoreConfiguration)
                else:
                    CableDict[key].append("Unknown")
                """
                All these for statements to make sure we get a proper ordening of attributes in the list:
                [Connection1, Coreconfiguration, Core Number, Connection2,
                 LSignal1..50, LTermination1..50, LSCR1..50, LCore1..50,
                 RCore1..50, RSCR1..50, RTermination1..50, RSignal1..50] 
                """
            # Next compile EnclosureDict
            # This contains the Enclosures with for each enclosure all cables attached to it.
            # As solarpanels also contain small JBs these are also added.
            # Format EnclosureTag: [ Cable 1, Cable2,....,SPARE,SPARE,..]
            if isinstance(obj,DS_BES_Enclosure) or isinstance(obj,DS_BES_Solarpanel):
                #print(obj.TagNumber)
                #print(obj.__dict__)
                EnclosureDict[key] = []
                print(key)
                for i in range(1,31):
                    attribute = 'Connection'+str(i)
                    if attribute in obj.__dict__:
                        #print(attribute,obj.__dict__[attribute])
                        if obj.__dict__[attribute] != "" or obj.__dict__[attribute] != None:
                            EnclosureDict[key].append(obj.__dict__[attribute])
                        #else:
                            #EnclosureDict[key].append("SPARE")
#        #print( "CableDict Length",len(CableDict[list(CableDict.keys())[0]]))
#        #print(CableDict[list(CableDict.keys())[0]])
        connection.close()
        storage.close()
        #print('EnclosureDict ', EnclosureDict)
#        # compile a connection Dictionary
#        # containing all cables per connection like so:
#        # Connection: [Cable1,Cable2,Cable3.....]
#        # e.g.: SPM-IJB-3510 ['IC-SPM-IJB-3510-01', 'IC-SPM-IJB-3510-02']
#        # So this construct can be used to obtain cable info for Cable1: 
#        # CableDict[ConnectionDict[0]]
#        ConnectionDict = {}
#        for cable in CableDict:
#            if CableDict[cable][0] not in ConnectionDict:
#                ConnectionDict[CableDict[cable][0]] = [] # Connection1 of the cable
#            ConnectionDict[CableDict[cable][0]].append(cable)
#            if len(CableDict[cable]) == 364: # contains all termination info
#                if CableDict[cable][3] not in ConnectionDict:
#                    ConnectionDict[CableDict[cable][3]] = []
#                ConnectionDict[CableDict[cable][3]].append(cable)
#            if len(CableDict[cable]) == 3: # doesn't contain termination info
#                if CableDict[cable][2] not in ConnectionDict:
#                    ConnectionDict[CableDict[cable][2]] = []
#                ConnectionDict[CableDict[cable][2]].append(cable)
##        for key in ConnectionDict:
#            #print(key, ConnectionDict[key])
#        CableDict = dict(sorted(CableDict.items()))
##        for key in CableDict:
##            print(key, CableDict[key])                
#    #######################################################################
#         #Export the Cable Dictionary to an Excel Workbook
        path = os.getcwd()+"\Terminations.xlsm"
        if os.path.exists(path):
            os.remove(path)
        time.sleep(2) # To give the os module time to remove the Excel Export File
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Terminations = Xcel.Workbooks.Add() 
        Terminations.SaveAs(path,52)    
        Xcel.Visible = True
        xlHAlignCenter = -4108
        Header2 = ["Signal","Term#","Core#","Configuration","Core#","Term#","Signal"]    
        row = 1
        # First cable end to end 
        for cable in CableDict:
            #print(cable)
            progress.update()
            startrow = row
            # make the Headers
            Terminations.ActiveSheet.Cells(row,1).Value = CableDict[cable][0] # Connection1
            Terminations.ActiveSheet.Range("A"+str(row)+":"+"B"+str(row)).Merge()
            Terminations.ActiveSheet.Cells(row,3).Value = cable # Cable Tag
            Terminations.ActiveSheet.Range("C"+str(row)+":"+"E"+str(row)).Merge()
            Terminations.ActiveSheet.Cells(row,6).Value = CableDict[cable][2] # Connection2
            Terminations.ActiveSheet.Range("F"+str(row)+":"+"G"+str(row)).Merge()
            row += 1
            Terminations.ActiveSheet.Cells(row,1).Value = Header2[0]
            Terminations.ActiveSheet.Cells(row,2).Value = Header2[1]
            Terminations.ActiveSheet.Cells(row,3).Value = Header2[2]
            Terminations.ActiveSheet.Cells(row,4).Value = CableDict[cable][-1] # coreconfiguration
            Terminations.ActiveSheet.Cells(row,5).Value = Header2[4]
            Terminations.ActiveSheet.Cells(row,6).Value = Header2[5]
            Terminations.ActiveSheet.Cells(row,7).Value = Header2[6]
            row += 1
            """
            [x, y, 'IS'], [x, y, 'G'], [x, y, 'OS'], [x, y, 'PE'] 
            x = number set s (pair, triad, quad), y = cores per set
            """
            if type(CableDict[cable][1]) is list:
                # Make Cell Names from Cable Name
                Name = "aa"+cable.replace('/','') # remove forwardslashes from cable names
                Name = Name.replace('-','') # remove dashes from cable names
                # the second entry in CableDict[cable] is a list [corenumber, core config (Pr,Tr), screen type]
                if CableDict[cable][1][2] == 'IS': # individual screen per pair or triad or quad
                    corenumber = CableDict[cable][1][0]*CableDict[cable][1][1] # total number of cores = corenumber x core config
                    multiplier = CableDict[cable][1][1] # core config
                    counter = 0
                    screennumber = 0
                    for i in range(1,corenumber+1):
                        progress.update()
                        LSignal = Name + "LSignal"+str(i)
                        #print(LSignal)
                        LTermination = Name + "LTermination" + str(i)                      
                        Color = Name + "Color" + str(i)                       
                        RTermination = Name + "RTermination" + str(i)
                        RSignal = Name + "RSignal"+str(i)
                        Terminations.ActiveSheet.Cells(row,1).Value = CableDict[cable][i+2]    # LSignal starts on index 3
                        Terminations.ActiveSheet.Cells(row,1).Name = LSignal
                        Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                        Terminations.ActiveSheet.Cells(row,2).Value = CableDict[cable][i+52]   # LTermination starts on index 53
                        Terminations.ActiveSheet.Cells(row,2).Name = LTermination
                        Terminations.ActiveSheet.Cells(row,3).Value = CableDict[cable][i+352] #  CoreNumber starts at index 353 
                        Terminations.ActiveSheet.Cells(row,4).Value = CableDict[cable][i+152]  # Color start at index 153
#                        Terminations.ActiveSheet.Cells(row,4).Name = Color
                        Terminations.ActiveSheet.Cells(row,5).Value = CableDict[cable][i+352] # CoreNumber starts at index 353 
                        Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"                        
                        Terminations.ActiveSheet.Cells(row,6).Value = CableDict[cable][i+252]  # RTermination starts on index 253
                        Terminations.ActiveSheet.Cells(row,6).Name = RTermination
                        Terminations.ActiveSheet.Cells(row,7).Value = CableDict[cable][i+302]  # RSignal starts on index 303
                        Terminations.ActiveSheet.Cells(row,7).Name = RSignal                            
                        row += 1
                        counter += 1
                        if counter == multiplier:
                            screennumber += 1
                            LSCR = Name + "LSCR" + str(screennumber)
                            RSCR = Name + "RSCR" + str(screennumber)
                            Terminations.ActiveSheet.Cells(row,2).Value = CableDict[cable][screennumber+102]   # LSCR starts on index 103
                            Terminations.ActiveSheet.Cells(row,2).Name = LSCR
                            Terminations.ActiveSheet.Cells(row,3).Value = "SCR"+str(screennumber)                            
                            Terminations.ActiveSheet.Cells(row,5).Value = "SCR"+str(screennumber)
                            Terminations.ActiveSheet.Cells(row,6).Value = CableDict[cable][screennumber+202]  # RSCR starts on index 203
                            Terminations.ActiveSheet.Cells(row,6).Name = RSCR
                            row +=1
                            counter = 0 # reset the multiplier counter
                if CableDict[cable][1][2] == 'OS' or CableDict[cable][1][2] == 'PE':
                    screenconfig = CableDict[cable][1][2]
                    corenumber = CableDict[cable][1][0]*CableDict[cable][1][1] # total number of cores
                    screennumber = 1
                    for i in range(1,corenumber+1):
                        progress.update()
                        LSignal = Name + "LSignal"+str(i)
                        #print(LSignal)
                        LTermination = Name + "LTermination" + str(i)                      
                        Color = Name + "Color" + str(i)                       
                        RTermination = Name + "RTermination" + str(i)
                        RSignal = Name + "RSignal"+str(i)
                        Terminations.ActiveSheet.Cells(row,1).Value = CableDict[cable][i+2]    # LSignal starts on index 3
                        Terminations.ActiveSheet.Cells(row,1).Name = LSignal
                        Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                        Terminations.ActiveSheet.Cells(row,2).Value = CableDict[cable][i+52]   # LTermination starts on index 53
                        Terminations.ActiveSheet.Cells(row,2).Name = LTermination
                        Terminations.ActiveSheet.Cells(row,3).Value = CableDict[cable][i+352] #  CoreNumber starts at index 353 
                        Terminations.ActiveSheet.Cells(row,4).Value = CableDict[cable][i+152]  # Color start at index 153
#                        Terminations.ActiveSheet.Cells(row,4).Name = Color
                        Terminations.ActiveSheet.Cells(row,5).Value = CableDict[cable][i+352] #  CoreNumber starts at index 353
                        Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"
                        Terminations.ActiveSheet.Cells(row,6).Value = CableDict[cable][i+252]  # RTermination starts on index 253
                        Terminations.ActiveSheet.Cells(row,6).Name = RTermination
                        Terminations.ActiveSheet.Cells(row,7).Value = CableDict[cable][i+302]  # RSignal starts on index 303
                        Terminations.ActiveSheet.Cells(row,7).Name = RSignal                            
                        row += 1
                    LSCR = Name + "LSCR" + str(screennumber)
                    RSCR = Name + "RSCR" + str(screennumber)
                    Terminations.ActiveSheet.Cells(row,2).Value = CableDict[cable][screennumber+102]   # LSCR starts on index 103
                    Terminations.ActiveSheet.Cells(row,2).Name = LSCR                    
                    Terminations.ActiveSheet.Cells(row,3).Value = screenconfig                            
                    Terminations.ActiveSheet.Cells(row,5).Value = screenconfig
                    Terminations.ActiveSheet.Cells(row,6).Value = CableDict[cable][screennumber+202]  # RSCR starts on index 203
                    Terminations.ActiveSheet.Cells(row,6).Name = RSCR
                    row +=1
            if type(CableDict[cable][1]) is str:
                Terminations.ActiveSheet.Cells(row,4).Value = CableDict[cable][1]
                row += 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(9).LineStyle = 1 # Continous line
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(7).LineStyle = 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(10).LineStyle = 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(9).Weight = 4 # Thick linestyle
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(7).Weight = 4
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(10).Weight = 4
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(12).LineStyle = 1 # internal cell borders
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(11).LineStyle = 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(12).Weight = 2 # internal cell borders Thin
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"G"+str(row-1)).Borders(11).Weight = 2

##############################################################################            
###        # Next per connection (Instrument or Enclosure)
        row = 1
        HandledCables = []
##        print(CableDict)
#        print(EnCab)
        for key in EnclosureDict:
            progress.update()
            startrow = row
            Terminations.ActiveSheet.Cells(row,9).Value = key
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"P"+str(row)).Merge() 
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"P"+str(row)).Interior.ColorIndex = 15 
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"P"+str(row)).Font.Bold = True
            row += 1
            for item in EnclosureDict[key]: # item is a Cable Tag or None
                progress.update()
                if item != None:
                    startrow2 = row
                    cable = item
                    Terminations.ActiveSheet.Cells(row,9).Value = Header2[0]
                    Terminations.ActiveSheet.Cells(row,10).Value = Header2[1]
                    Terminations.ActiveSheet.Cells(row,11).Value = Header2[2]
                    Terminations.ActiveSheet.Cells(row,12).Value = cable # cable tag
                    Terminations.ActiveSheet.Cells(row,12).Interior.ColorIndex = 4
                    Terminations.ActiveSheet.Cells(row,13).Value = Header2[4]
                    Terminations.ActiveSheet.Cells(row,14).Value = Header2[5]
                    Terminations.ActiveSheet.Cells(row,15).Value = Header2[6]
                    condition1 = CableDict[item][2] != key # Enclosure is on left side of the cable
                    condition2 = CableDict[item][2] == key # Enclosure is on right side of the cable
                    if condition1:
                        Terminations.ActiveSheet.Cells(row,16).Value = CableDict[cable][2]
                    if condition2:
                        Terminations.ActiveSheet.Cells(row,16).Value = CableDict[cable][0]
                    row += 1
                    if item not in HandledCables:
                        prefix = "bb"
                    else:
                        prefix = "cc"                
                    if type(CableDict[item][1]) is list: # now always the case
                        Name = prefix+item.replace('/','') # remove forwardslashes from cable names
                        Name = Name.replace('-','') # remove dashes from cable names
                        if CableDict[item][1][2] == 'IS': # individual screen per pair or triad or quad
                            corenumber = CableDict[item][1][0]*CableDict[item][1][1] # total number of cores
                            multiplier = CableDict[item][1][1]
                            counter = 0
                            screennumber = 0
                            for i in range(1,corenumber+1):
                                progress.update()
                                LSignal = Name + "LSignal"+str(i)
                                #print(LSignal)
                                LTermination = Name + "LTermination" + str(i)                      
                                Color = Name + "Color" + str(i)                       
                                RTermination = Name + "RTermination" + str(i)
                                RSignal = Name + "RSignal" + str(i)
                                if condition1:
                                    Terminations.ActiveSheet.Cells(row,9).Value = CableDict[item][i+2]    # LSignal starts on index 3
                                    Terminations.ActiveSheet.Cells(row,9).Name = LSignal                               
                                if condition2:
                                    Terminations.ActiveSheet.Cells(row,9).Value = CableDict[item][i+302]  # RSignal starts on index 302
                                    Terminations.ActiveSheet.Cells(row,9).Name = RSignal   
                                if condition1:
                                    Terminations.ActiveSheet.Range("J"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,10).Name = LTermination
                                if condition2:
                                    Terminations.ActiveSheet.Range("J"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][i+252]   # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,10).Name = RTermination                                    
                                Terminations.ActiveSheet.Cells(row,11).Value = CableDict[item][i+352] #  CoreNumber starts at index 353 
                                Terminations.ActiveSheet.Cells(row,12).Value = CableDict[item][i+152]  # Color start at index 153
#                                Terminations.ActiveSheet.Cells(row,12).Name = Color
                                Terminations.ActiveSheet.Cells(row,13).Value = CableDict[item][i+352] # CoreNumber starts at index 353
                                if condition1:
                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,14).Name = RTermination
                                if condition2:
                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+52]  # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,14).Name = LTermination
                                if condition1:
                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+302]  # RSignal starts on index 303
                                    Terminations.ActiveSheet.Cells(row,15).Name = RSignal  
                                if condition2:
                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+2]  # LSignal starts on index 3
                                    Terminations.ActiveSheet.Cells(row,15).Name = LSignal                                    
                                row += 1
                                counter += 1
                                if counter == multiplier:
                                    screennumber += 1
                                    LSCR = Name + "LSCR" + str(screennumber)
                                    RSCR = Name + "RSCR" + str(screennumber)
                                    if condition1:
                                        Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][screennumber+102]   # LSCR starts on index 103
                                        Terminations.ActiveSheet.Cells(row,10).Name = LSCR
                                    if condition2:
                                        Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][screennumber+202]   # RSCR starts on index 203
                                        Terminations.ActiveSheet.Cells(row,10).Name = RSCR                                        
                                    Terminations.ActiveSheet.Cells(row,11).Value = "SCR"+str(screennumber)                            
                                    Terminations.ActiveSheet.Cells(row,13).Value = "SCR"+str(screennumber)
                                    if condition1:
                                        Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
                                        Terminations.ActiveSheet.Cells(row,14).Name = RSCR
                                    if condition2:
                                        Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+102]  # LSCR starts on index 103
                                        Terminations.ActiveSheet.Cells(row,14).Name = LSCR                                        
                                    row +=1
                                    counter = 0 # reset the multiplier counter
                        if CableDict[item][1][2] == 'OS' or CableDict[item][1][2] == 'PE':
                            screenconfig = CableDict[item][1][2]
                            corenumber = CableDict[item][1][0]*CableDict[item][1][1] # total number of cores
                            screennumber = 1
                            for i in range(1,corenumber+1):
                                progress.update()
                                LSignal = Name + "LSignal"+str(i)
                                #print(LSignal)
                                LTermination = Name + "LTermination" + str(i)                      
                                Color = Name + "Color" + str(i)                       
                                RTermination = Name + "RTermination" + str(i)
                                RSignal = Name + "RSignal"+ str(i)
                                if condition1:
                                    Terminations.ActiveSheet.Cells(row,9).Value = CableDict[item][i+2]    # LSignal starts on index 3
                                    Terminations.ActiveSheet.Cells(row,9).Name = LSignal
                                if condition2:
                                    Terminations.ActiveSheet.Cells(row,9).Value = CableDict[item][i+302]    # RSignal starts on index 302
                                    Terminations.ActiveSheet.Cells(row,9).Name = RSignal
                                if condition1:
                                    Terminations.ActiveSheet.Range("J"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,10).Name = LTermination
                                if condition2:
                                    Terminations.ActiveSheet.Range("J"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][i+252]   # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,10).Name = RTermination                                    
                                Terminations.ActiveSheet.Cells(row,11).Value = CableDict[item][i+352] #  CoreNumber starts at index 353 
                                Terminations.ActiveSheet.Cells(row,12).Value = CableDict[item][i+152]  # Color start at index 153
#                                Terminations.ActiveSheet.Cells(row,12).Name = Color
                                Terminations.ActiveSheet.Cells(row,13).Value = CableDict[item][i+352] #  CoreNumber starts at index 353
                                if condition1:
                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,14).Name = RTermination
                                if condition2:
                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+52]  # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,14).Name = LTermination
                                if condition1:                                    
                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+302]  # RSignal starts on index 303
                                    Terminations.ActiveSheet.Cells(row,15).Name = RSignal
                                if condition2:
                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+2]  # LSignal starts on index 3
                                    Terminations.ActiveSheet.Cells(row,15).Name = LSignal                                   
                                row += 1
                            LSCR = Name + "LSCR" + str(screennumber)
                            RSCR = Name + "RSCR" + str(screennumber)
                            if condition1:
                                Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][screennumber+102]   # LSCR starts on index 103
                                Terminations.ActiveSheet.Cells(row,10).Name = LSCR
                            if condition2:
                                Terminations.ActiveSheet.Cells(row,10).Value = CableDict[item][screennumber+202]   # RSCR starts on index 203
                                Terminations.ActiveSheet.Cells(row,10).Name = RSCR                                
                            Terminations.ActiveSheet.Cells(row,11).Value = screenconfig                            
                            Terminations.ActiveSheet.Cells(row,13).Value = screenconfig
                            if condition1:
                                Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
                                Terminations.ActiveSheet.Cells(row,14).Name = RSCR
                            if condition2:
                                Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+102]  # LSCR starts on index 103
                                Terminations.ActiveSheet.Cells(row,14).Name = LSCR                                
                            row +=1
                    if type(CableDict[item][1]) is str:
                        Terminations.ActiveSheet.Cells(row,12).Value = CableDict[item][1]
                        row += 1
                    HandledCables.append(item)
                    Terminations.ActiveSheet.Range("P"+str(startrow2)+":"+"P"+str(row-1)).Merge()
                    Terminations.ActiveSheet.Range("P"+str(startrow2)+":"+"P"+str(row-1)).VerticalAlignment = -4160 # Align top
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(9).LineStyle = 1 # Continous line
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(7).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(10).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(9).Weight = 4 # Thick linestyle
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(7).Weight = 4
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(10).Weight = 4
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(12).LineStyle = 1 # internal cell borders
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(11).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(12).Weight = 2 # internal cell borders Thin
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"P"+str(row-1)).Borders(11).Weight = 2
        Terminations.ActiveSheet.Range("A:P").Columns.AutoFit()
        Terminations.ActiveSheet.Range("A:P").HorizontalAlignment = xlHAlignCenter            
        # Add the macro for automatic cell changing
        xlButtonControl = 0 
        pathToMacro = os.getcwd()+"\Macro.txt"
        MacroName = "Button1_Click"
        with open (pathToMacro, "r") as myfile:
            print('reading macro into string from: ' + str(myfile))
            macro=myfile.read()
        excelModule = Terminations.VBProject.VBComponents.Add(1)
        excelModule.CodeModule.AddFromString(macro)
        UpdateButton = Terminations.ActiveSheet.Shapes.AddFormControl(xlButtonControl,539.25,15.75,43.5,33.75)
        UpdateButton.OnAction = "Button1_Click"
        Terminations.ActiveSheet.Shapes("Button 1").Left = Terminations.ActiveSheet.Range("H1").Left 
        Terminations.ActiveSheet.Shapes("Button 1").Top = Terminations.ActiveSheet.Range("H1").Top
        Terminations.ActiveSheet.Shapes("Button 1").TextFrame.Characters(1,8).Text = ""
        Terminations.ActiveSheet.Shapes("Button 1").TextFrame.Characters(1,1).Text = "UPDATE"
        Xcel.Application.Run(MacroName)    
        progress.stop()
        Terminations.Close(SaveChanges=True)
        Xcel.Application.Quit()
        tk.messagebox.showwarning(title=None, message="Terminations Export finished.")
        btn_terminationexport.config(state=tk.NORMAL)
        btn_terminationimport.config(state=tk.NORMAL)             
    except BaseException as e:
        print(e.args)
        btn_terminationexport.config(state=tk.NORMAL)
        btn_terminationimport.config(state=tk.NORMAL)        
        connection.close()
        storage.close()        
        progress.stop()     