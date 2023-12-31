import tkinter as tk
from tkinter import filedialog
from tkinter import simpledialog
import tkinter.messagebox
import tkinter.ttk as ttk
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
addr = '10.175.13.199', 8090 # This is the address of the ZEO Server
DS_List = os.listdir('Datasheets')
#DB_List = os.listdir('Databases') # for later use
exportprocess = None
PolyLineDict = {}
LineDict = {}
TagDict = {}
EnclosureDict = {}
CableDict = {}
CableTypeDict = {}
LoopDict = {}
HalfLoopDict = {}
# print(DS_List)
Datasheet = None
Object = None
RETAIN = False # Boolean needed to check whether an exisitng object was chosen from the Database
SELECTED = False # Boolean needed to check whether a category of objects has been chosen.
                 # used in functions objects and fillobject 
OBJ_List = []
Tag_List = []
OBJNumbersDict = {}
chosenTag = ""
Min = None
Max = None

Spares_List = ["2 Spares","4 Spares","6 Spares","8 Spares","10 Spares","12 Spares","14 Spares","16 Spares",
               "18 Spares","20 Spares","22 Spares","24 Spares","26 Spares","28 Spares","30 Spares"]
Spares_Terms = [2,4,6,8,10,12,14,16,18,20,22,24,26,28,30]
crawlcount = 0

# Initiate the Object Number Dictionary
for name, obj in inspect.getmembers(sys.modules[__name__]):
    if inspect.isclass(obj) and obj.__name__.startswith('DS_BES'):
        OBJNumbersDict[obj.__name__] = 0
# Fill OBJ_List with Objects from the Database
#directory = os.getcwd()+"/Databases"
#filename = "/Lufeng.fs"
# In this manner the database and helpfiles can be placed anywhere
# as long as the directory structure is maintained
#storage = FileStorage.FileStorage(directory+filename)
storage = ClientStorage.ClientStorage(addr)
db = DB(storage)        
connection = db.open()
root = connection.root()
for key in root:
    OBJ_List.append(key)
    for item in OBJNumbersDict:
        obj = root[key]
        if item == type(root[key]).__name__:
            OBJNumbersDict[item] += 1 
connection.close()
storage.close()

# Fill CableTypeDict with data from the CableTypes.xlsx
# The build up of the Dicitionary is as follows:
# key: [Manufacturer,Fire Protection,Core Configuration,Voltage Rating,Diameter Area,Details,
#       Outer Diameter, Glandsize, Rated Current, Resistance, Capacitance, Unit Mass, Unit Cost,
#       Cores per Multiplier , Multiplier (Pair, Triad etc.), Screen Type]
print("Filling Cable Type Dictionary")
directory = os.getcwd()+"/PF Datasheets"
filename = "/CableTypes.xlsx"
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Xcel.Visible = False
CableTypeSheet = Xcel.Workbooks.Open(directory+filename)
xlUp = -4162
LastRow = CableTypeSheet.ActiveSheet.Cells(CableTypeSheet.ActiveSheet.Rows.Count, "A").End(xlUp).Row
for row in range(2,LastRow+1):
    print('\r'+str(round((row/LastRow)*100))+'%', end='\r')
    ValueList = []
    for column in range(2,15):
        ValueList.append(CableTypeSheet.ActiveSheet.Cells(row,column).Value)
    CableTypeDict[CableTypeSheet.ActiveSheet.Cells(row,1).Value] =ValueList
    row += 1
CableTypeSheet.Close(SaveChanges=False)
Xcel.Application.Quit()

for key in CableTypeDict:
    coreconfig = CableTypeDict[key][2]   
    if coreconfig.find('Pr')>=0:
        CableTypeDict[key].append(int(coreconfig[:coreconfig.find('Pr')]))
        CableTypeDict[key].append(2) # 2 cores per set (Pair)     
    if coreconfig.find('Tr')>=0:
        CableTypeDict[key].append(int(coreconfig[:coreconfig.find('Tr')]))
        CableTypeDict[key].append(3) # 3 cores per set (Triad)
    if coreconfig.find('IS')>=0:
        CableTypeDict[key].append('IS')
    if coreconfig.find('OS')>=0:
        CableTypeDict[key].append('OS')
    if coreconfig.find('x')>=0:
        if coreconfig.find('G')>=0:
            CableTypeDict[key].append(int(coreconfig[:coreconfig.find('G')])) # E cables with separate Ground core
            CableTypeDict[key].append(1) # multiplier is always 1
            CableTypeDict[key].append('G')
        else:
            CableTypeDict[key].append(int(coreconfig[:coreconfig.find('x')])) # E cables
            CableTypeDict[key].append(1) # multiplier is always 1
            CableTypeDict[key].append('PE') # E cables apparently have no IS or OS      
#print(CableTypeDict)
#Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
#Xcel.Visible = True
#Xcel = Xcel.Workbooks.Add()
#column = 1
#row = 1
#for key in CableTypeDict:
#    Xcel.ActiveSheet.Cells(row,column).Value =  key
#    Xcel.ActiveSheet.Cells(row,column+1).Value = str(CableTypeDict[key])
#    Xcel.ActiveSheet.Cells(row,column+2).Value = str(CableTypeDict[key][-3:])
#    column=1
#    row += 1  
#print(CableTypeDict[key])       
print("CableTypeDict filled")
## Fill the Core Type Dicitonary
print("Filling Core Type Dictionary")
CoreTypeDict = {}
directory = os.getcwd()+"/PF Datasheets"
filename = "/Core Type.xlsx"
Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
Xcel.Visible = False
CoreTypeSheet = Xcel.Workbooks.Open(directory+filename)
xlUp = -4162
LastRow = CoreTypeSheet.ActiveSheet.Cells(CoreTypeSheet.ActiveSheet.Rows.Count, "A").End(xlUp).Row
for row in range(2,LastRow+1):
    print('\r'+str(round((row/LastRow)*100))+'%', end='\r')
    if CoreTypeSheet.ActiveSheet.Cells(row,1).Value not in CoreTypeDict:
        CoreTypeDict[CoreTypeSheet.ActiveSheet.Cells(row,1).Value] = []
    ValueTuple = (CoreTypeSheet.ActiveSheet.Cells(row,2).Value,
                  CoreTypeSheet.ActiveSheet.Cells(row,3).Value,
                  CoreTypeSheet.ActiveSheet.Cells(row,4).Value)
    CoreTypeDict[CoreTypeSheet.ActiveSheet.Cells(row,1).Value].append(ValueTuple)
CoreTypeSheet.Close(SaveChanges=False)
Xcel.Application.Quit()
print("CoreTypeDict Filled")
#print(CoreTypeDict)

#Apparently for importing Excel date fields this statement is necessary
pywintypes.datetime = pywintypes.TimeType

# auxiliary function to return RGB for Excel cell coloring
def rgbToInt(rgb):
    colorInt = rgb[0] + (rgb[1] * 256) + (rgb[2] * 256 * 256)
    return colorInt
 
def select(): # Button btn_diagram clicked
    print("SELECT") #TODO add to logfile
    directory = os.getcwd()+"/ACAD"
    selectedFile = filedialog.askopenfilename(initialdir = directory,
                                                 title = "Select file",
                                                 filetypes = (("acad files","*.dwg"),("all files","*.*")))
    if selectedFile == "": # cancel button clicked
        return
    filename = selectedFile[len(directory)+1:]
    print("filename",filename)
    lbl_diagram["text"] = filename # used as transport to function crawl()
    
def startcrawl():
    print("STARTCRAWL")
    global crawlcount
    crawlcount += 1
    exec("crawlthread"+str(crawlcount)+"= threading.Thread(target=crawl).start()")
        
def crawl(): # Button btn_crawl clicked
    print("CRAWL")
    pythoncom.CoInitialize()
    global Tag_List, crawlthread
    Tag_List = []
    OriginalTag_ListLength = len(Tag_List)
    # empty the listbox
    lbox_tags.delete(0, tk.END)
    TxtIPointDict = {}
    AttrIPointDict = {}
    try:
        directory = os.getcwd()+"/ACAD"
        # check whether an acad file has been selected
        filename = lbl_diagram["text"] # set in function select()
        # disable the buttons to prevent failures
        btn_diagram.config(state=tk.DISABLED)
        btn_crawl.config(state=tk.DISABLED)
        btn_tagexport.config(state=tk.DISABLED)
        btn_connections.config(state=tk.DISABLED)
        btn_loopexport.config(state=tk.DISABLED)
        if filename == "Selected Diagram": # no acad file selected
            tk.messagebox.showwarning(title=None, message="First select a Diagram.")
            # enable the buttons
            btn_diagram.config(state=tk.NORMAL)
            btn_crawl.config(state=tk.NORMAL)
            btn_tagexport.config(state=tk.NORMAL)
            btn_connections.config(state=tk.NORMAL)
            btn_loopexport.config(state=tk.NORMAL)
            return
        else:
            filename = "/" + lbl_diagram["text"]
        print(directory+filename)
        acad = win32com.client.Dispatch("Autocad.Application")
        doc = acad.Documents.Open(directory+filename)
        acad.Visible = False
        progress['mode'] = 'indeterminate'
        progress.start()
        progress.update()
        time.sleep(2)
        #----------------------- AcDbText Handling ------------------------------------------------ 
        # Apparently there are F&G layaouts where Attributes are combined with Text to form a tag number.
        # In order to obtain the correct text by the Attributes, it is required to first collect the 
        # Insertion Points of the Text entities. If these form a tag by themselves, they obviously contain
        # a dash: "-"  
        counter = 0
        Tag = ''
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.ObjectName         
            if name == 'AcDbText': # try finding tags in Text objects (Single Line)
                counter += 1
                match0 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+-[\d\w/]+', entity.TextString)
                match1 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match2 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match3 = re.search(r'[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match4 = re.search(r'[\d\w]+-[\d\w]+', entity.TextString)
                if match0 and match1 and match2 and match3:                    # for cable numbers IC-SPM-HPU-9401-SV03-01
                    #print(match1.group())
                    Tag_List.append(match0.group())
                    TagDict[match0.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match1.group()].append(entity.InsertionPoint)
                if not(match0) and match1 and match2 and match3:               # for cable numbers like C-110-SOV-1003-O
                    #print(match1.group())
                    Tag_List.append(match1.group())
                    TagDict[match1.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match1.group()].append(entity.InsertionPoint)
                elif not(match1) and match2 and match3: # for cable numbers like C-110-PD-020               
#                    print(match2.group())
                    Tag_List.append(match2.group())
                    TagDict[match2.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match2.group()].append(entity.InsertionPoint)                
                elif not(match1) and not(match2) and match3:
                    #print(match3.group())
                    Tag_List.append(match3.group())
                    TagDict[match3.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match3.group()].append(entity.InsertionPoint)
                elif not(match1) and not(match2) and not(match3) and match4:
                    #print(match4.group())
                    Tag_List.append(match4.group())
                    TagDict[match4.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match4.group()].append(entity.InsertionPoint)                      
        #----------------------- AcDbMText Handling ------------------------------------------------ 
            if name == 'AcDbMText': # try finding tags in MText objects (Multi Line)
                Multitext = entity.TextString
                string1 = re.sub(r'\\.*?;',r' ', Multitext) # remove odd characters
                string2 = re.sub(r'\\P',r' ', string1) # remove odd characters
                match1 = re.findall(r'[\d\w]+-[\d\w]+-[\d\w]+', string2) # for normal tags like 110-WT-1001B1            
                match2 = re.findall(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+', string2) # for cable numbers like C-110-PD-020
                if match2 and match1:
                    for item1 in match1:
                        for item2 in match2:
                            if item1 in item2:
                                match1.remove(item1)
                if match2:
                    for item in match2:
                        match1.append(item)                                
                for item in match1:
#                    print(item)
                    Tag_List.append(item)
                    #TagDict[Tag] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[Tag].append(entity.InsertionPoint)
        #----------------------- AcDbBlockReference Handling ------------------------------------------------                              
#            print(name, entity.Handle) #TODO add to logfile 
#        made changes to include the LMI RECTANGLE
            if name == 'AcDbBlockReference':
                    print(entity.name, entity.layer)
                    HasAttributes = entity.HasAttributes
                    if entity.name == "LMI RECTANGLE" and entity.layer == "Dim":
                        if HasAttributes:  
                            for attrib in entity.GetAttributes():  
                                if attrib.TagString == "IF":
                                    IF = attrib.TextString
                                if attrib.TagString == "TN":
                                    TN = attrib.TextString
                                if attrib.TagString == "SC":
                                    SC = attrib.TextString
                            if IF and TN and SC: 
                                Tag = SC+'-'+IF+'-'+TN
                                Tag_List.append(Tag)
                               
                    IF, TN, SC = '', '', '' 
                    if HasAttributes:  
                        for attrib in entity.GetAttributes():
#                            print(attrib.TagString)
#                            print(attrib.TextString)
                            # Sometimes a complete tag is in one attribute usually recognizabele
                            # because it contains a dash
                            if "-" in attrib.Textstring:
                                Tag_List.append(attrib.TextString)
#                                print(Tag_List)
                                break
                            if attrib.TagString == "IF":
                                IF = attrib.TextString
                            if attrib.TagString == "TN":
                                TN = attrib.TextString
                            if attrib.TagString == "SC":
                                SC = attrib.TextString
                        if IF and TN and SC:       
                            Tag = SC+'-'+IF+'-'+TN
#                            print(Tag)
                            Tag_List.append(Tag)
                        if IF and TN and not SC:
#                            print(Tag)
                            Tag = IF+'-'+TN
                            Tag_List.append(Tag)
                        if not IF and not TN and not SC:                               
                            # for other attributes we need further investigation per attribute
                            # hence all attributes are first collected in a list
                            AttribList = []
                            for attrib in entity.GetAttributes():
    #                            print(attrib.TagString)
    #                            print(attrib.TextString)                         
                                AttribList.append(attrib.TextString)
                            for item in AttribList:
                                if item in InstrumentFunction:
#                                    print(AttribList)
                                    position = AttribList.index(item)
                                    if position == 0: 
                                        # no system number added to instrument tag
                                        #start searching for a system number
                                        SystemNumber = "NOSYS#"
                                        for tag in Tag_List:
                                            StringList = tag.split("-")
                                            if len(StringList) == 3:
                                                SystemNumber = StringList[0]
                                                break
                                            if len(StringList) > 4:
                                                SystemNumber = StringList[2]
                                                break
                                        Tag = SystemNumber+"-"+AttribList[position]+"-"+AttribList[position+1]
                                    else:
                                        Tag = AttribList[position-1]+"-"+AttribList[position]+"-"+AttribList[position+1]
                                    Tag_List.append(Tag)
            print(Tag_List)                        
            # remove any duplicates
            Tag_List = list(dict.fromkeys(Tag_List))
            if len(Tag_List) > OriginalTag_ListLength:
                OriginalTag_ListLength = len(Tag_List)
                lbox_tags.delete(0, tk.END)
                # lbox_tags.update_idletasks()
                Tag_List.sort()
                for item in Tag_List:
                    lbox_tags.insert(tk.END, item)
                    lbox_tags.yview(tk.END)
                    lbox_tags.update_idletasks()
            progress.update()
            time.sleep(0.2)            
        #----------------------- AcDbAttributeDefinition Handling ------------------------------------------------                                 
        print("ATTRIBUTE")
        AttrIPointDict1 = {}
        nameList = []
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.ObjectName
            if name not in nameList:
                nameList.append(name)
            #print(name, entity.Handle) #TODO add to logfile 
            if name == 'AcDbAttributeDefinition': # in the HB AutoCad layouts Attribute definitions are used
                AttrIPointDict[entity.Handle] = (entity.InsertionPoint,entity.TagString)
                AttrIPointDict1[entity.Handle] = (entity.InsertionPoint,entity.TagString)
                # InsertionPoint = (x,y,z)
#                AttrHandle.append(entity.Handle)
            progress.update()
        print(nameList)               
        for key in AttrIPointDict:           
            xdistance = 0
            ydistance = 0
            AttrLst = []
            xkey = AttrIPointDict[key][0][0]
            ykey = AttrIPointDict[key][0][1]
            AttrLst.append((key,AttrIPointDict[key][1]))
            for key1 in AttrIPointDict1:
                if key != key1:
                    xkey1 = AttrIPointDict1[key1][0][0]
                    ykey1 = AttrIPointDict1[key1][0][1]
                    xdistance = abs(xkey - xkey1)
                    ydistance = abs(ykey - ykey1)
                    if xdistance < 1400 and ydistance < 800:
                        AttrLst.append((key1,AttrIPointDict1[key1][1]))
#                        print(AttrLst)
                    # AttrLst = [(Handle1,Tag1),(Handle2,Tag2),(Handle3,Tag3)]
                progress.update()
            if len(AttrLst) == 3:
                    Tag = ''
                    xpoint1 = AttrIPointDict[AttrLst[0][0]][0][0]
                    xpoint2 = AttrIPointDict[AttrLst[1][0]][0][0]
                    xpoint3 = AttrIPointDict[AttrLst[2][0]][0][0]                    
                    ypoint1 = AttrIPointDict[AttrLst[0][0]][0][1]
                    ypoint2 = AttrIPointDict[AttrLst[1][0]][0][1]
                    ypoint3 = AttrIPointDict[AttrLst[2][0]][0][1]
                    lowest_ypoint = 0.0
                    lowest_xpoint = 0.0
                    text_ypoint = 0.0
                    text_xpoint = 0.0
                    condition1 = ypoint1 > ypoint2 > ypoint3
                    condition2 = ypoint1 > ypoint3 > ypoint2
                    condition3 = ypoint2 > ypoint1 > ypoint3
                    condition4 = ypoint2 > ypoint3 > ypoint1
                    condition5 = ypoint3 > ypoint1 > ypoint2
                    condition6 = ypoint3 > ypoint2 > ypoint1
                    if condition1:
                        Tag = AttrLst[0][1]+"-"+AttrLst[1][1]+"-"+AttrLst[2][1]
                        lowest_ypoint = ypoint3
                        lowest_xpoint = xpoint3
                    if condition2:
                        Tag = AttrLst[0][1]+"-"+AttrLst[2][1]+"-"+AttrLst[1][1]
                        lowest_ypoint = ypoint2
                        lowest_xpoint = xpoint2                        
                    if condition3:
                        Tag = AttrLst[1][1]+"-"+AttrLst[0][1]+"-"+AttrLst[2][1]
                        lowest_ypoint = ypoint3
                        lowest_xpoint = xpoint3                        
                    if condition4:
                        Tag = AttrLst[1][1]+"-"+AttrLst[2][1]+"-"+AttrLst[0][1]
                        lowest_ypoint = ypoint1
                        lowest_xpoint = xpoint1                        
                    if condition5:
                        Tag = AttrLst[2][1]+"-"+AttrLst[0][1]+"-"+AttrLst[1][1]
                        lowest_ypoint = ypoint2
                        lowest_xpoint = xpoint2                        
                    if condition6:
                        Tag = AttrLst[2][1]+"-"+AttrLst[1][1]+"-"+AttrLst[0][1] 
                        lowest_ypoint = ypoint1
                        lowest_xpoint = xpoint1                      
                    for text in TxtIPointDict: # Search for the corresponding text to complete the tag.                           
                        text_xpoint = TxtIPointDict[text][0][0] 
                        text_ypoint = TxtIPointDict[text][0][1]
                        condition7 = abs(lowest_xpoint - text_xpoint) < 300.0 and abs(lowest_ypoint - text_ypoint)< 1000.0
                        if condition7:
                            Tag = Tag+"-"+TxtIPointDict[text][1]
                    if Tag not in Tag_List:
                        Tag_List.append(Tag)
                    AttrLst = []
            progress.update()
            print(Tag_List)
            # only update listbox if the Tag_List has expended                          
            if len(Tag_List) > OriginalTag_ListLength:
                OriginalTag_ListLength = len(Tag_List)
                lbox_tags.delete(0, tk.END)
                # lbox_tags.update_idletasks()
                Tag_List.sort()
                for item in Tag_List:
                    lbox_tags.insert(tk.END, item)
                    lbox_tags.yview(tk.END)
                    lbox_tags.update_idletasks()
            time.sleep(0.5)
        # print(Tag_List)
#        print(TxtIPointDict)
#        print(AttrIPointDict)        
        doc.Close(True)
        # Remove the Acad application made previously
        os.system('TASKKILL /F /IM acad.exe')
        progress.stop()
        tk.messagebox.showwarning(title=None, message="Examination Finished.")
        pythoncom.CoUninitialize()
        btn_diagram.config(state=tk.NORMAL)
        btn_crawl.config(state=tk.NORMAL)
        btn_tagexport.config(state=tk.NORMAL)
        btn_connections.config(state=tk.NORMAL)
        btn_loopexport.config(state=tk.NORMAL)
    except BaseException as e:
        print(e.args)
        tk.messagebox.showerror(title=None, message=e.args)
        # doc.Close()
        # Remove the Acad application made previously
        pythoncom.CoUninitialize()
        btn_diagram.config(state=tk.NORMAL)
        btn_crawl.config(state=tk.NORMAL)
        btn_tagexport.config(state=tk.NORMAL)
        btn_connections.config(state=tk.NORMAL)
        btn_loopexport.config(state=tk.NORMAL)
        progress.stop()
        os.system('TASKKILL /F /IM acad.exe')
        
def tagexport():
    print("TAGEXPORT")
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()    
    tags_tuple = lbox_tags.get(0,last=tk.END)
    if len(tags_tuple) == 0:
        progress.stop()
        tk.messagebox.showwarning(title=None, message="No tags\nexamined yet.")
        return        
#    print(tags_tuple)
#    for name, obj in inspect.getmembers(sys.modules[__name__]):
#        if inspect.isclass(obj):
#            print(obj.__name__)    
    path = os.getcwd()+"\Tags Export.xlsx"
    # Remove exisiting collection and create new one
    if os.path.exists(path):
        os.remove(path)
    time.sleep(2) # To give the os module time to remove the Excel Export File    
    Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    Export = Xcel.Workbooks.Add() 
    Export.SaveAs(path)
    # Add the headers
    Export.ActiveSheet.Cells(1,1).Value = "TAG"
    Export.ActiveSheet.Cells(1,1).Font.Bold = True
    Export.ActiveSheet.Cells(1,2).Value = "                  OBJECT                  "
    Export.ActiveSheet.Cells(1,2).Font.Bold = True
    Export.ActiveSheet.Cells(1,3).Value = "DIAGRAM"
    Export.ActiveSheet.Cells(1,3).Font.Bold = True    
    Export.ActiveSheet.Cells(2,3).Value = lbl_diagram["text"]   
    row = 2
    for tag in tags_tuple:
        Export.ActiveSheet.Cells(row,1).Value = tag
        row += 1
        progress.update()
    counter = 10000
    for name, obj in inspect.getmembers(sys.modules[__name__]):
        if inspect.isclass(obj) and obj.__name__.startswith("DS"):
            Export.ActiveSheet.Cells(counter,108).Value = obj.__name__
            counter += 1
    xlUp = -4162
    LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "DD").End(xlUp).Row
    Range = "=$DD$10000:$DD$" + str(LastRow)
#    for i in range(1,100):
#        Cell = "B"+str(i)
#        if Export.ActiveSheet.Range(Cell).Validation.InCellDropdown == True:
#            Export.ActiveSheet.Range(Cell).Validation.Delete()
    LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    for i in range(1,LastRow+1):
        Cell = "B"+str(i)
        Export.ActiveSheet.Range(Cell).Validation.Add(3, 1, 1, Range)
        Export.ActiveSheet.Range(Cell).Validation.InCellDropdown = True            
    Export.ActiveSheet.Range("A:EE").Columns.AutoFit()
    Export.Close(SaveChanges=True)        
    Xcel.Application.Quit()
    lbl_diagram["text"] = "Selected Diagram"    
    progress.stop()
    tk.messagebox.showwarning(title=None, message="Tags Export finished.")
    os.system('TASKKILL /F /IM excel.exe') 
    
def startconnections():
    print("STARTINDEXIMPORTATION")
    threading.Thread(target=connections).start()    

def connections():     
    print("CONNECTIONS")
    global CableDict,LineDict,EnclosureDict,PolyLineDict,TagDict,LoopDict,HalfLoopDict,Tag_List
    Tag_List = []
    Min = None
    Max = None
    OriginalTag_DictLength = 0
    try:
        btn_diagram.config(state=tk.DISABLED)
        btn_crawl.config(state=tk.DISABLED)
        btn_tagexport.config(state=tk.DISABLED)
        btn_connections.config(state=tk.DISABLED)
        btn_loopexport.config(state=tk.DISABLED)
        progress['mode'] = 'indeterminate'
        progress.start()
        progress.update()
        directory = os.getcwd()+"/ACAD"
        filename = lbl_diagram["text"] # set in function select()
        if filename == "Selected Diagram": # no acad file selected
            progress.stop()
            tk.messagebox.showwarning(title=None, message="First select a Diagram.")
            btn_diagram.config(state=tk.NORMAL)
            btn_crawl.config(state=tk.NORMAL)
            btn_tagexport.config(state=tk.NORMAL)
            btn_connections.config(state=tk.NORMAL)
            btn_loopexport.config(state=tk.NORMAL)
            return
        else:
            filename = "/" + lbl_diagram["text"] 
        pythoncom.CoInitialize()
        acad = win32com.client.Dispatch("Autocad.Application")
        doc = acad.Documents.Open(directory+filename)
        acad.Visible = False
        time.sleep(2)
        # drawings in ModelSpace have to be drawn within a certain matrix
        # only entities within the matrix have to be investigated for connections
        Min, Max = None, None
        minX, minY, maxX, maxY = 0.0, 0.0, 0.0, 0.0
        # matrix = doc.Blocks.Item('Matrix') #TODO
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.ObjectName
            # print(entity,name)
            if name == 'AcDbBlockReference' and entity.Name == 'Matrix':
               maxX, maxY = entity.GetBoundingBox(Min,Max)[1][0], entity.GetBoundingBox(Min,Max)[1][1]
               minX, minY = entity.InsertionPoint[0], entity.InsertionPoint[1]
               print(minX, minY)
               print(maxX,maxY)
        time.sleep(1)                        
        Flag = False
        # Check whether the diagram to be examined is actually a Blockdiagram by
        # checking for 'Blockdiagram" or 'Block Diagram'etc. in the title
        for entity in acad.ActiveDocument.PaperSpace:
            name = entity.ObjectName
            if Flag == False and name == 'AcDbBlockReference':
                HasAttributes = entity.HasAttributes
                if HasAttributes:
                    for attrib in entity.GetAttributes():
                        if attrib.TextString:
                            # remove all whitespaces from the TextString
                            TString = attrib.TextString.translate({ord(c): None for c in string.whitespace})
                            TString.capitalize()
                            if "BLOCKDIAGRAM" in TString:
                                Flag = True
            if Flag == False and name == 'AcDbText':
                TString = entity.TextString.translate({ord(c): None for c in string.whitespace})
                TString.capitalize()
                if "BLOCKDIAGRAM" in TString:
                    Flag = True           
            if Flag == False and name == 'AcDbMText':
                TString = entity.TextString.translate({ord(c): None for c in string.whitespace})
                TString.capitalize()
                if "BLOCKDIAGRAM" in TString:
                    Flag = True
        if Flag == False:
            tk.messagebox.showwarning(title=None, message="This is apparently not\na Block Diagram")
            progress.stop()
            doc.close()
            acad.Application.Quit()
            del acad
            if "DADispatcherService.exe" in (p.name() for p in psutil.process_iter()):
                os.system("taskkill /f /im  DADispatcherService.exe") 
            btn_diagram.config(state=tk.NORMAL)
            btn_crawl.config(state=tk.NORMAL)
            btn_tagexport.config(state=tk.NORMAL)
            btn_connections.config(state=tk.NORMAL)
            btn_loopexport.config(state=tk.NORMAL)
            return            
        time.sleep(3)
        counter = 0
        # First collect Tags
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.ObjectName
            # print(name,entity.Handle)
    #----------------------------------------------------------------------------------------------------        
            if name == 'AcDbText': # try finding tags in Text objects (Single Line)
                match0 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+-[\d\w/]+', entity.TextString)
                match1 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match2 = re.search(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match3 = re.search(r'[\d\w]+-[\d\w]+-[\d\w/]+', entity.TextString)
                match4 = re.search(r'[\d\w]+-[\d\w]+', entity.TextString)
                if match0 and match1 and match2 and match3:                    # for cable numbers IC-SPM-HPU-9401-SV03-01
                    #print(match1.group())
                    Tag_List.append(match0.group())
                    TagDict[match0.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match1.group()].append(entity.InsertionPoint)
                if not(match0) and match1 and match2 and match3:               # for cable numbers like C-110-SOV-1003-O
                    #print(match1.group())
                    Tag_List.append(match1.group())
                    TagDict[match1.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match1.group()].append(entity.InsertionPoint)
                elif not(match1) and match2 and match3: # for cable numbers like C-110-PD-020               
                    #print(match2.group())
                    Tag_List.append(match2.group())
                    TagDict[match2.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match2.group()].append(entity.InsertionPoint)                
                elif not(match1) and not(match2) and match3:
                    #print(match3.group())
                    Tag_List.append(match3.group())
                    TagDict[match3.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match3.group()].append(entity.InsertionPoint)
                elif not(match1) and not(match2) and not(match3) and match4:
                    #print(match4.group())
                    Tag_List.append(match4.group())
                    TagDict[match4.group()] = list(entity.GetBoundingBox(Min,Max))
                    #TagDict[match4.group()].append(entity.InsertionPoint)                
    #----------------------------------------------------------------------------------------------------
            if name == 'AcDbMText': # try finding tags in MText objects (Multi Line)
                Multitext = entity.TextString
                #print(repr(Multitext))
                string1 = re.sub(r'\\.*?;',r' ', Multitext) # remove odd characters
                string2 = re.sub(r'\\P',r' ', string1) # remove odd characters
                match1 = re.findall(r'[\d\w]+-[\d\w]+-[\d\w]+', string2) # for normal tags like 110-WT-1001B1            
                match2 = re.findall(r'[\d\w]+-[\d\w]+-[\d\w]+-[\d\w]+', string2) # for cable numbers like C-110-PD-020
                if match2 and match1:
                    #print(match1,match2)
                    for item1 in match1:
                        for item2 in match2:
                            if item1 in item2:
                                match1.remove(item1)
                if match2:
                    for item in match2:
                        match1.append(item)                                
                for item in match1:
                    #print(item)
                    Tag_List.append(item)
                    TagDict[item] = list(entity.GetBoundingBox(Min,Max))
    ##----------------------------------------------------------------------------------------------------
            if name == 'AcDbBlockReference': # try finding tags in BlockReference objects
                IF, TN, SC = '', '', '' 
                print('Block:', entity.Handle)
                HasAttributes = entity.HasAttributes
                if HasAttributes:
                    for attrib in entity.GetAttributes():
                        print(attrib.TagString, attrib.TextString)
                        # Sometimes a complete tag is in one attribute usually recognizabele
                        # because it contains a dash                    
                        # if "-" in attrib.Textstring:
                        #    Tag_List.append(attrib.TextString)
                        #    TagDict[attrib.TextString] = list(entity.GetBoundingBox(Min,Max))
                            #TagDict[attrib.TextString].append(entity.InsertionPoint)
                        #    break
                        # Normal BlockReference has TagStrings: IF, TN and SC
                        if attrib.TagString == "IF":
                            IF = attrib.TextString
                        if attrib.TagString == "TN":
                            TN = attrib.TextString
                        if attrib.TagString == "SC":
                            SC = attrib.TextString                     
                    if IF and TN and SC:       
                        Tag = SC+'-'+IF+'-'+TN
                        print('3-Tagstring',Tag)
                        Tag_List.append(Tag)
                        TagDict[Tag] = list(entity.GetBoundingBox(Min,Max))
                    if IF and TN and not SC:
                        print('2-Tagstring',Tag)
                        Tag = IF+'-'+TN
                        Tag_List.append(Tag)
                        TagDict[Tag] = list(entity.GetBoundingBox(Min,Max))
                    if not IF and not TN and not SC:
                        AttribList = []
                        for attrib in entity.GetAttributes():               
                            AttribList.append(attrib.TextString)
                        for item in AttribList:
                            if item in InstrumentFunction:
                                # print(AttribList)
                                position = AttribList.index(item)
                                if position == 0: # probably no system number added to instrument tag
                                    #start searching for a system number
                                    SystemNumber = "NOSYS#"
                                    for tag in Tag_List:
                                        StringList = tag.split("-")
                                        if len(StringList) == 3:
                                            SystemNumber = StringList[0]
                                            break
                                        if len(StringList) > 4:
                                            SystemNumber = StringList[2]
                                            break
                                    Tag = SystemNumber+"-"+AttribList[position]+"-"+AttribList[position+1]
                                else:
                                    Tag = AttribList[position-1]+"-"+AttribList[position]+"-"+AttribList[position+1]
                                print('1-Tagstring',Tag)
                                Tag_List.append(Tag)
                                TagDict[Tag] = list(entity.GetBoundingBox(Min,Max))
                                #TagDict[Tag].append(entity.InsertionPoint)
                if len(TagDict) > OriginalTag_DictLength:
                    OriginalTag_DictLength = len(TagDict)
                    lbox_tags.delete(0, tk.END)
                    # lbox_tags.update_idletasks()
                    Tag_List.sort()
                    for item in Tag_List:
                        lbox_tags.insert(tk.END, item)
                        lbox_tags.yview(tk.END)
                        lbox_tags.update_idletasks()
                time.sleep(0.5)
        counter += 1
        #print(counter)                 
        time.sleep(0.2) # to give ACAD time to respond
    ##=========================================== GRAPHBUILDING========================================
    ##--------------------------Search Enclosures -----------------------------------------------------
        print("IN GRAPHBUILDING")
#        print(Tag_List)
        progress.update()
        for entity in acad.ActiveDocument.ModelSpace:
            name = entity.ObjectName
            key = entity.Handle                      
            if name == 'AcDbPolyline': # collect Polylines as possible enclosures
                # check whether the Polyline is closed and not dashed and doesn't have
                # LinetypeGeneration enabled   
                condition = (entity.Closed) and not ("DASH" in entity.Linetype) and (entity.LinetypeGeneration == False) 
                if condition: 
                    Min = None
                    Max = None
                    #print( entity.GetBoundingBox(Min,Max))
                    PolyLineDict[key] = entity.GetBoundingBox(Min,Max)
            counter += 1
            #print(counter)
            if name == 'AcDbLine':
                if not "DASH" in entity.Linetype:
                    Min = None
                    Max = None
                    LineDict[key] = entity.GetBoundingBox(Min,Max) # used in cable search
            time.sleep(0.1)
        #print(TagDict)
        #print(LineDict)
        def contained(taglist,polylinelist):
            # this function checks whether a tag is surrounded by
            # a closed polyline. This is regarded as a enclosure
            x1 = taglist[0][0]
            y1 = taglist[0][1]
            x2 = taglist[1][0]
            y2 = taglist[1][1]
            x3 = polylinelist[0][0]
            y3 = polylinelist[0][1]
            x4 = polylinelist[1][0]
            y4 = polylinelist[1][1]
            xc = x1 + 0.5*(x2-x1) # calculate centerpoint x-coordinate
            yc = y1 + 0.5*(y2-y1) # calculate centerpoint y-coordinate
            if x3<xc<x4 and y3<yc<y4:
                return True
            else:
                return False
        def vicinity(taglist,linelist):
            # this function checks whether a tag is in the vicinity of
            # a line. In that case the line is regarded as a cable
            x1 = taglist[0][0] # Min
            y1 = taglist[0][1]
            x2 = taglist[1][0] # Max
            y2 = taglist[1][1]
            x3 = linelist[0][0] # Min
            y3 = linelist[0][1]
            x4 = linelist[1][0] # Max
            y4 = linelist[1][1]
            condition1 = abs(x2-x1)/abs(y2-y1) #>1 tag horizontal <1 tag vertical
            condition2 = abs(x4-x3)<2 # line vertical
            condition3 = abs(y4-y3)<2 # line horizontal
            # vertical because in that case abs(x1-x2)/abs(y1-y2) >= 1        
            if y3<y1<y2<y4 and condition1<1 and condition2:
                if abs(x2-x3)<2.5:
                    return True
            # horizontal because in that case abs(x1-x2)/abs(y1-y2) < 1
            if x3<x1<x2<x4 and condition1>1 and condition3:
                if abs(y1-y3)<2.5:
                    return True
            else:
                return False   
        def attached(object1,cable,distance):
            # this function checks whether a line is connected to 
            # a polyline recognized as an enclosure. In that case
            # it is taken as a cable-enclosure connection.
            x1 = object1[0][0]
            y1 = object1[0][1]
            x2 = object1[1][0]
            y2 = object1[1][1]
            xs = cable[0][0]
            ys = cable[0][1]
            xe = cable[1][0]
            ye = cable[1][1]
            if abs(xs-x2)<distance and y1<ys<y2: #Horizontal left
                return True
            if abs(xe-x1)<distance and y1<ye<y2: #Horizontal right
                return True
            if abs(ys-y2)<2 and x1<xs<x2: #Vertical bottom
                return True        
            if abs(ye-y1)<2 and x1<xe<x2: #Vertical top
                return True
            else:
                return False
        def connected(line1,line2):
            # this function checks whether 2 lines are connected at
            # their starting or endpoints in which case it is regarded
            # as the proceeding of a cable.
            xs1 = line1[0][0]
            ys1 = line1[0][1]
            xe1 = line1[1][0]
            ye1 = line1[1][1]
            xs2 = line2[0][0]
            ys2 = line2[0][1]
            xe2 = line2[1][0]
            ye2 = line2[1][1]
            if abs(xe1-xs2)<2 and abs(ye1-ys2)<2: #Horizontal left
                return True
            if abs(xe1-xe2)<2 and abs(ye1-ye2)<2: #Horizontal right
                return True
            if abs(xs1-xs2)<2 and abs(ys1-ys2)<2: #Vertical bottom
                return True        
            if abs(xs1-xe2)<2 and abs(ys1-ye2)<2: #Vertical top
                return True
            if abs(xe1-xs2)<2 and ys2<ye1<ye2:
                return True
            if abs(xs1-xe2)<2 and ys2<ye1<ye2:
                return True
            if abs(xe1-xe2)<2 and ys2<ye1<ye2:
                return True
            if abs(xs1-xs2)<2 and ys2<ye1<ye2:
                return True        
            else:
                return False        
        for tag in TagDict:
            for polyline in PolyLineDict:
                if contained(TagDict[tag],PolyLineDict[polyline]):
                    EnclosureDict[tag] = polyline
#        print(TagDict)
#        print(EnclosureDict)
        for tag in TagDict:        
            for line in LineDict:
                if vicinity(TagDict[tag],LineDict[line]):
                    CableDict[tag] = line
#                    print(CableDict)
#        print(EnclosureDict)
#        print(CableDict)
        # try to detect whether there are extra characters in the cable tag
        Diff =[]
        for tag in TagDict:
            for cable in CableDict:
                if (tag in cable) and len(cable)>len(tag) and not(cable.split("-")[0] in Diff):
                    Diff.append(cable.split("-")[0])
                    break
#        print(Diff)            
        CableList = []
        # get a list of tags that are apparently not cables  and remove these          
        for cable in list(CableDict.keys()):
            indicator = False
            for element in Diff:
                if element in cable:
                    indicator = True
            if indicator == False:
                CableList.append(cable)
#        print(CableList)
        for cable in CableList:
            del CableDict[cable]
        # Loop compiling
        print("Zeroth round")      
        for cable in CableDict:
            LoopList = []
            Ctuple = LineDict[CableDict[cable]]
#            print(Ctuple)
            for enclosure in EnclosureDict:
                Etuple = PolyLineDict[EnclosureDict[enclosure]] 
                if attached(Etuple,Ctuple,2):
                    HalfLoopDict[cable] = [Ctuple]
                    LoopList.append(enclosure)
                    #print(cable,enclosure)
            for tag in TagDict:
#                print(TagDict)
                if tag not in CableDict and tag not in EnclosureDict:
                    Ttuple = TagDict[tag]
                    if attached(Ttuple,Ctuple,7):
                        HalfLoopDict[cable] = [Ctuple]
                        LoopList.append(tag)
                        print(cable,tag)
            LoopDict[cable] = LoopList
        # remove cables that are connected after the first round from HalfLoopDict
        print(LoopDict)
        for cable in LoopDict:
            if len(LoopDict[cable])==2 and cable in HalfLoopDict:
                del HalfLoopDict[cable]        
        # Start searching for cables that have only one connection
    #===========================================================
        def lineconnections():
            global CableDict,LineDict,EnclosureDict,PolyLineDict,TagDict,LoopDict,HalfLoopDict        
            print("First round") 
            for cable in LoopDict:
                if len(LoopDict[cable])<2: # cables under one angle
                    # get the Line Handle from CableDict
                    # print(cable, LoopDict[cable])
                    LineTuple1 = HalfLoopDict[cable][-1]
                    for line in LineDict:
                        LineTuple2 = LineDict[line]               
                        if (LineTuple1 != LineTuple2) and connected(LineTuple1,LineTuple2):
                            HalfLoopDict[cable].append(LineTuple2)
                            for enclosure in EnclosureDict:
                                Etuple = PolyLineDict[EnclosureDict[enclosure]] 
                                if attached(Etuple,LineTuple2,2):
                                    LoopDict[cable].append(enclosure)
                                    #print(cable,enclosure)
                            for tag in TagDict:
                                if tag not in CableDict and tag not in EnclosureDict:
                                    Ttuple = TagDict[tag]
                                    if attached(Ttuple,LineTuple2,7):
                                        LoopDict[cable].append(tag)                               
                                        #print(cable,tag)
            # remove cables that are connected after the first round from HalfLoopDict
            for cable in LoopDict:
                if len(LoopDict[cable])==2 and cable in HalfLoopDict:
                    del HalfLoopDict[cable]                        
            print("Second round")                                
            for cable in LoopDict:
                if len(LoopDict[cable])<2: # cables under two angles
                    # print(cable, LoopDict[cable])
                    LineTuple1 = HalfLoopDict[cable][-1]
                    for line in LineDict:
                        LineTuple2 = LineDict[line]  
                        if (LineTuple1 != LineTuple2) and connected(LineTuple1,LineTuple2):
                            if not(LineTuple2 in HalfLoopDict[cable]): # this is were the code d
                                HalfLoopDict[cable].append(LineTuple2)
                                for enclosure in EnclosureDict:
                                    Etuple = PolyLineDict[EnclosureDict[enclosure]] 
                                    if attached(Etuple,LineTuple2,2):
                                        if not (enclosure in LoopDict[cable]):
                                            LoopDict[cable].append(enclosure)
                                        #print(cable,enclosure)
                                for tag in TagDict:
                                    if tag not in CableDict and tag not in EnclosureDict:
                                        Ttuple = TagDict[tag]
                                        if attached(Ttuple,LineTuple2,7):
                                            if not(tag in LoopDict[cable]):
                                                LoopDict[cable].append(tag)
                                            #print(cable,tag)
            # remove cables that are connected after the second round from HalfLoopDict
            for cable in LoopDict:
                if len(LoopDict[cable])==2 and cable in HalfLoopDict:
                    del HalfLoopDict[cable]                                
            print("Third round")                                
            for cable in LoopDict:
                if len(LoopDict[cable])<2: # cables under two angles
                    LineTuple1 = LineDict[CableDict[cable]]            
                    for line in LineDict:
                        LineTuple2 = HalfLoopDict[cable][-1]              
                        if (LineTuple1 != LineTuple2) and connected(LineTuple1,LineTuple2):
                            if not(LineTuple2 in HalfLoopDict[cable]):
                                HalfLoopDict[cable].append(LineTuple2)                    
                                for enclosure in EnclosureDict:
                                    Etuple = PolyLineDict[EnclosureDict[enclosure]] 
                                    if attached(Etuple,LineTuple2,2):
                                        if not (enclosure in LoopDict[cable]):
                                            LoopDict[cable].append(enclosure)
                                        #print(cable,enclosure)
                                for tag in TagDict:
                                    if tag not in CableDict and tag not in EnclosureDict:
                                        Ttuple = TagDict[tag]
                                        if attached(Ttuple,LineTuple2,7):
                                            if not(tag in LoopDict[cable]):
                                                LoopDict[cable].append(tag)
                                            #print(cable,tag)
            # remove cables that are connected after the third round from HalfLoopDict                                   
            for cable in LoopDict:
                if len(LoopDict[cable])==2 and cable in HalfLoopDict:
                    del HalfLoopDict[cable]
    #==================================================================
        lineconnections()                
        # find interrupted lines
        # print(HalfLoopDict)
        def endline(lineTuple1,lineTuple2):
            xs1 = lineTuple1[0][0]
            ys1 = lineTuple1[0][1]
            xe1 = lineTuple1[1][0]
            ye1 = lineTuple1[1][1]
            xs2 = lineTuple2[0][0]
            ys2 = lineTuple2[0][1]
            xe2 = lineTuple2[1][0]
            ye2 = lineTuple2[1][1]
            if abs(xe1-xs2)<2 and abs(ye1-ys2)<2 and xs1<xs2 and ye1<ye2: # 1
                return ["Horizontal Left",xs1,ys1]
            if abs(xe1-xe2)<2 and abs(ye1-ye2)<2 and xs1<xe2 and ys2<ye1: # 2
                return ["Horizontal Left",xs1,ys1]
            if abs(xs1-xs2)<2 and abs(ys1-ys2)<2 and xs2<xe1 and ye1<ye2: # 3
                return ["Horizontal right",xe1,ye1]
            if abs(xs1-xe2)<2 and abs(ys1-ye2)<2 and xe2<xe1 and ys2<ys1: # 4
                return ["Horizontal right",xe1,ye1]
            if abs(xs1-xs2)<2 and abs(ys1-ys2)<2 and xe1<xe2 and ys2<ye1: # 5
                return ["Vertical Top",xe1,ye1]
            if abs(xs1-xe2)<2 and abs(ys1-ye2)<2 and xs2<xs1 and ye2<ye1: # 6
                return ["Vertical Top",xe1,ye1]
            if abs(xe1-xe2)<2 and abs(ye1-ye2)<2 and xs2<xe1 and ys1<ye2: # 7
                return ["Vertical Bottom",xs1,ys1]
            if abs(xe1-xs2)<2 and abs(ye1-ys2)<2 and xe1<xe2 and ys1<ye2: # 8
                return ["Vertical Bottom",xs1,ys1]
            else:
                return "Empty"
            
        if len(HalfLoopDict)>0:
            for cable in HalfLoopDict:
                if len(HalfLoopDict[cable])>1:
                    lineTuple1 = HalfLoopDict[cable][-1]
                    lineTuple2 = HalfLoopDict[cable][-2]
                    print(cable,endline(lineTuple1,lineTuple2))
                    L = endline(lineTuple1,lineTuple2)
                    if L[0] == "Horizontal Left":
                        for line in LineDict:
                            if LineDict[line][0][0]<LineDict[line][1][0]<L[1] and\
                            abs(L[2]-LineDict[line][0][1])<2.0:
                                HalfLoopDict[cable].append(LineDict[line])
                    if L[0] == "Horizontal right":
                        for line in LineDict:
                            if L[1]<LineDict[line][0][0]<LineDict[line][1][0] and\
                            abs(L[2]-LineDict[line][0][1])<2.0:
                                HalfLoopDict[cable].append(LineDict[line])                    
                    if L[0] == "Vertical Top":
                        for line in LineDict:
                            if L[2]<LineDict[line][0][1]<LineDict[line][1][1] and\
                            abs(L[1]-LineDict[line][0][0])<2.0:
                                HalfLoopDict[cable].append(LineDict[line])                      
                    if L[0] == "Vertical Bottom":
                        for line in LineDict:
                           if LineDict[line][0][1]<LineDict[line][1][1]<L[2] and\
                           abs(L[1]-LineDict[line][0][0])<2.0:
                                HalfLoopDict[cable].append(LineDict[line])
                if len(HalfLoopDict[cable])==1:
#                    print(cable,HalfLoopDict[cable])
                    for line in LineDict:
                        lineTuple1 = HalfLoopDict[cable][-1]
                        lineTuple2 = LineDict[line]
                        if not(lineTuple2 in HalfLoopDict[cable]):
                            xs1 = lineTuple1[0][0]
                            ys1 = lineTuple1[0][1]
                            xe1 = lineTuple1[1][0]
                            ye1 = lineTuple1[1][1]
                            xs2 = lineTuple2[0][0]
                            ys2 = lineTuple2[0][1]
                            xe2 = lineTuple2[1][0]
                            ye2 = lineTuple2[1][1]                        
                            # find lines with (nearly) the same y coordinates
                            condition = False
                            if abs(ys1-ys2)<2.0:
                                if abs(xs1-xs2)<30.0 or abs(xs1-xe2)<30.0 or abs(xe1-xs2)<30.0 or abs(xe1-xe2)<30.0:
                                    condition = True
    #                        # find lines with (nearly) the same x coordinates
    #                        if abs(xs1-xs2)<2.0:
    #                            if abs(ys1-ys2)<30.0 or abs(ys1-ye2)<30.0 or abs(ye1-ys2)<30.0 or abs(ye1-ye2)<30.0:
    #                                condition = True
                            if condition:
                                HalfLoopDict[cable].append(LineDict[line])
#                    print(cable,HalfLoopDict[cable])                            
        for cable in HalfLoopDict:
            Ctuple = HalfLoopDict[cable][-1]
            for enclosure in EnclosureDict:
                Etuple = PolyLineDict[EnclosureDict[enclosure]] 
                if attached(Etuple,Ctuple,2):
                    LoopDict[cable].append(enclosure)
                    #print(cable,enclosure)
            for tag in TagDict:
                if tag not in CableDict and tag not in EnclosureDict:
                    Ttuple = TagDict[tag]
                    if attached(Ttuple,Ctuple,7):
                        LoopDict[cable].append(tag)
                        #print(cable,tag)
        # remove cables that are connected from HalfLoopDict
        for cable in LoopDict:
            if len(LoopDict[cable])==2 and cable in HalfLoopDict:
                del HalfLoopDict[cable]                                    
        lineconnections()
        for cable in LoopDict:
            if len(LoopDict[cable])==2:
                s2 = LoopDict[cable][0]
                if s2 in cable:
                    # swap source and destination
                    LoopDict[cable][0],LoopDict[cable][1] = LoopDict[cable][1],LoopDict[cable][0]
        print(LoopDict)                                      
        doc.close()
        acad.Application.Quit()
        del acad
        if "DADispatcherService.exe" in (p.name() for p in psutil.process_iter()):
            os.system("taskkill /f /im  DADispatcherService.exe")
        progress.stop()
        lbl_diagram["text"] = "Selected Diagram"
        tk.messagebox.showwarning(title=None, message="Examination Finished.")
        pythoncom.CoUninitialize()
        btn_diagram.config(state=tk.NORMAL)
        btn_crawl.config(state=tk.NORMAL)
        btn_tagexport.config(state=tk.NORMAL)
        btn_connections.config(state=tk.NORMAL)
        btn_loopexport.config(state=tk.NORMAL)
        return
    except BaseException as e:
        pythoncom.CoUninitialize()
        btn_diagram.config(state=tk.NORMAL)
        btn_crawl.config(state=tk.NORMAL)
        btn_tagexport.config(state=tk.NORMAL)
        btn_connections.config(state=tk.NORMAL)
        btn_loopexport.config(state=tk.NORMAL)
        print(e.args)
        #doc.close()
        progress.stop()
        acad.Application.Quit()
        if "DADispatcherService.exe" in (p.name() for p in psutil.process_iter()):
            os.system("taskkill /f /im  DADispatcherService.exe")
            
def loopexport():
    print("LOOPEXPORT")
    global LoopDict
    if len(LoopDict) == 0:
        tk.messagebox.showwarning(title=None, message="No connections\nexamined yet.")
        return
    progress['mode'] = 'indeterminate'
    progress.start()   
    progress.update()    
    path = os.getcwd()+"\Connections Export.xlsx"
    # Remove exisiting collection and create new one
    if os.path.exists(path):
        os.remove(path)
    time.sleep(2) # To give the os module time to remove the Excel Export File    
    Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    Export = Xcel.Workbooks.Add() 
    Export.SaveAs(path)
    # Add the headers
    Export.ActiveSheet.Cells(1,1).Value = "CABLE"
    Export.ActiveSheet.Cells(1,1).Font.Bold = True
    Export.ActiveSheet.Cells(1,2).Value = "CONNECTION1"
    Export.ActiveSheet.Cells(1,2).Font.Bold = True
    Export.ActiveSheet.Cells(1,3).Value = "CONNECTION2"
    Export.ActiveSheet.Cells(1,3).Font.Bold = True      
    row = 2
    for cable in LoopDict:
        if len(LoopDict[cable]) == 2:
            Export.ActiveSheet.Cells(row,1).Value = cable
            Export.ActiveSheet.Cells(row,2).Value = LoopDict[cable][0]
            Export.ActiveSheet.Cells(row,3).Value = LoopDict[cable][1]
            row += 1
            progress.update()
        else:
            Export.ActiveSheet.Cells(row,1).Value = cable
            Export.ActiveSheet.Cells(row,2).Value = LoopDict[cable][0]
            Export.ActiveSheet.Cells(row,3).Value = "N/A"
            row += 1
            progress.update()
    # Sorting of enclosures to one side
    xlUp = -4162
    LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "A").End(xlUp).Row
    Export.ActiveSheet.Range("A:EE").Columns.AutoFit()
    Export.Close(SaveChanges=True)        
    Xcel.Application.Quit()
    progress.stop()
    tk.messagebox.showwarning(title=None, message="Connections Export\nfinished.")
    os.system('TASKKILL /F /IM excel.exe')             
    
def new(tag):
    print("NEW")
    global Datasheet, Object, chosenTag, OBJ_List, DS_List, addr
    print(tag)
    New_DS = lbox_datasheets.curselection() # Returns a tuple 
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and not New_DS:
            tk.messagebox.showwarning(title=None, message="First select a Datasheet.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    print(DS_List[New_DS[0]]) # First number of tuple determines position in DS_list
    if DS_List[New_DS[0]] == "DS_BES_AI.xlsx":
        Object = DS_BES_AI()
        EmptySheet = "DS_BES_AI.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_AO.xlsx":
        Object = DS_BES_AO()
        EmptySheet = "DS_BES_AO.xlsx"     
    if DS_List[New_DS[0]] == "DS_BES_Antenna.xlsx":        
        Object = DS_BES_Antenna()
        EmptySheet = "DS_BES_Antenna.xlsx"    
    if DS_List[New_DS[0]] == "DS_BES_Beacon.xlsx":
        Object = DS_BES_Beacon()
        EmptySheet = "DS_BES_Beacon.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Cable.xlsx":
        Object = DS_BES_Cable()
        EmptySheet = "DS_BES_Cable.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Compass.xlsx":
        Object = DS_BES_Compass()
        EmptySheet = "DS_BES_Cable.xlsx"        
    if DS_List[New_DS[0]] == "DS_BES_CV.xlsx":
        Object = DS_BES_CV()
        EmptySheet = "DS_BES_CV.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_DI.xlsx":
        Object = DS_BES_DI()
        EmptySheet = "DS_BES_DI.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_DO.xlsx":
        Object = DS_BES_DO()
        EmptySheet = "DS_BES_DO.xlsx"        
    if DS_List[New_DS[0]] == "DS_BES_Enclosure.xlsx":
        Object = DS_BES_Enclosure()
        EmptySheet = "DS_BES_Enclosure.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_FD.xlsx":
        Object = DS_BES_FD()
        EmptySheet = "DS_BES_FD.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_FI.xlsx":
        Object = DS_BES_FI()
        EmptySheet = "DS_BES_FI.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Fogdetector.xlsx":
        Object = DS_BES_Fogdetector()
        EmptySheet = "DS_BES_Fogdetector.xlsx" 
    if DS_List[New_DS[0]] == "DS_BES_Foghorn.xlsx":
        Object = DS_BES_Foghorn()
        EmptySheet = "DS_BES_Foghorn.xlsx"         
    if DS_List[New_DS[0]] == "DS_BES_GD.xlsx":
        Object = DS_BES_GD()
        EmptySheet = "DS_BES_GD.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Handsw.xlsx":
        Object = DS_BES_Handsw()
        EmptySheet = "DS_BES_Handsw.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_LG.xlsx":
        Object = DS_BES_LG()
        EmptySheet = "DS_BES_LG.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Limitsw.xlsx":
        Object = DS_BES_Limitsw()
        EmptySheet = "DS_BES_Limitsw.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_LIT.xlsx":
        Object = DS_BES_LIT()
        EmptySheet = "DS_BES_LIT.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Loadpin.xlsx":
        Object = DS_BES_Loadpin()
        EmptySheet = "DS_BES_Loadpin.xlsx" 
    if DS_List[New_DS[0]] == "DS_BES_Oceanograph.xlsx":
        Object = DS_BES_Oceanograph()
        EmptySheet = "DS_BES_Oceanograph.xlsx"         
    if DS_List[New_DS[0]] == "DS_BES_PG.xlsx":
        Object = DS_BES_PG()
        EmptySheet = "DS_BES_PG.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_PIG.xlsx":
        Object = DS_BES_PIG()        
        EmptySheet = "DS_BES_PIG.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_PIT.xlsx":
        Object = DS_BES_PIT()
        EmptySheet = "DS_BES_PIT.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_RO.xlsx":
        Object = DS_BES_RO()
        EmptySheet = "DS_BES_RO.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_SOL_V.xlsx":
        Object = DS_BES_SOL_V()
        EmptySheet = "DS_BES_SOL_V.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Solarpanel.xlsx":
        Object = DS_BES_Solarpanel()
        EmptySheet = "DS_BES_Solarpanel.xlsx" 
    if DS_List[New_DS[0]] == "DS_BES_Slipring.xlsx":
        Object = DS_BES_Slipring()
        EmptySheet = "DS_BES_Slipring.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Speaker.xlsx":
        Object = DS_BES_Speaker()
        EmptySheet = "DS_BES_Speaker.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_SV.xlsx":
        Object = DS_BES_SV()
        EmptySheet = "DS_BES_SV.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_TG.xlsx":
        Object = DS_BES_TG()
        EmptySheet = "DS_BES_TG.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_TT.xlsx":
        Object = DS_BES_TT()
        EmptySheet = "DS_BES_TT.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Weatherstation.xlsx":
        Object = DS_BES_Weatherstation()
        EmptySheet = "DS_BES_Weatherstation.xlsx"
    if DS_List[New_DS[0]] == "DS_BES_Windgenerator.xlsx":
        Object = DS_BES_Windgenerator()
        EmptySheet = "DS_BES_Windgenerator.xlsx"        
    if DS_List[New_DS[0]] == "DS_BES_Transformer.xlsx":
        Object = DS_BES_Transformer()
        EmptySheet = "DS_BES_Transformer.xlsx"        
    print(Object)            
    if DS_List[New_DS[0]].startswith("PF_"): # A pre filled datasheet has been selected
        print("Pre-filled selected")
        path = os.getcwd()+"\Datasheets\\"
        path1 = path + DS_List[New_DS[0]] # The prefilled datasheet
        path2 = path + EmptySheet # The blank datasheet
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        PF_Datasheet = Xcel.Workbooks.Open(path1)
        Datasheet = Xcel.Workbooks.Open(path2)
        for key in Object.InternalFieldsDict:
            Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value = \
            PF_Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value
        for key in Object.FieldsDict:
            Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value = \
            PF_Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value
        PF_Datasheet.Close(SaveChanges=False)
    else:
        path = os.getcwd()+"\Datasheets\\"
        path = path + DS_List[New_DS[0]]
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Datasheet = Xcel.Workbooks.Open(path) # to open the selected Datasheet
    if chosenTag != "No chosen Tag": # tagname has been chosen from the P&ID or Blockdiagram
        if chosenTag in OBJ_List:
            chosenTag = "No chosen Tag"
            lbl_chosenTag["text"] = chosenTag
            tk.messagebox.showwarning(title=None, message="Object is already in the Database.")
            Datasheet.Close(True)
            Datasheet = None
            return
        else:
            Datasheet.ActiveSheet.Range("_TAG_NO").Value = chosenTag # Fill in the Tag Number 
            Datasheet.ActiveSheet.Range("_PID_NO").Value = lbl_diagram["text"]
            # Fill in the P&ID Number
            chosenTag = "No chosen Tag"
            lbl_chosenTag["text"] = chosenTag 
#---------------------------------------------------------------------------------------------
# prepare validation lists of cables and of instruments for use in drop-down list        
    InstrumentsCount = 1
    CableCount = 1
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    InstrumentsValidateList = []
    CableValidateList = []
    for key in root:
        obj = root[key]
        if not isinstance(obj,DS_BES_Cable):
            InstrumentsValidateList.append(key)
            InstrumentsCount += 1
        if isinstance(obj,DS_BES_Cable):
            CableValidateList.append(key)
            CableCount += 1
    connection.close()
    storage.close()        
#---------------------------------------------------------------------------------------------
    ObjectName = DS_List[New_DS[0]]        
    if  ObjectName == "DS_BES_Cable.xlsx":
        print("Cable chosen in new")
        # To provide a pull down list in the datasheet of instruments or enclosures that the cables can connect to
        Range = "=$DD$2:$DD$" + str(len(InstrumentsValidateList)+1)
        counter = 2
        if InstrumentsValidateList:
            InstrumentsValidateList.sort()
            for item in InstrumentsValidateList:
                Datasheet.ActiveSheet.Cells(counter,108).Value = item # 108 is column DD
                counter += 1
            # The following commands are apparently necessary to insert an updated drop-down list in Excel
            if Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.Delete()
            if Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.Delete()
            # Reconstruction of the updated drop-down list
            Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.InCellDropdown = True
            Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.InCellDropdown = True    
#---------------------------------------------------------------------------------------------          
    # THe following instruments have either no cable or more than 2 cables 
    condition = ObjectName == "DS_BES_CV.xlsx" or ObjectName == "DS_BES_Enclosure.xlsx" or\
                ObjectName == "DS_BES_LG.xlsx" or ObjectName == "DS_BES_PG.xlsx" or\
                ObjectName == "DS_BES_TG.xlsx" or ObjectName == "DS_BES_RO.xlsx" or\
                ObjectName == "DS_BES_SV.xlsx" or ObjectName == "DS_BES_Cable.xlsx" or\
                ObjectName == "DS_BES_Transformer.xlsx" or  ObjectName == "DS_BES_Slipring.xlsx"
    if  not (condition):
        # To provide a pull down list in the datasheet of cables the instruments can connect to 
        print("Instrument chosen in new")
        Range = "=$DD$2:$DD$" + str(len(CableValidateList)+1)        
        counter = 2
        CableValidateList.sort()
        if CableValidateList:
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,108).Value = item
                counter += 1
            if Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_Connection1").Validation.Delete()
            if Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_Connection2").Validation.Delete()
            Datasheet.ActiveSheet.Range("_Connection1").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown = True
            Datasheet.ActiveSheet.Range("_Connection2").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown = True
#---------------------------------------------------------------------------------------------
    if  ObjectName == "DS_BES_Enclosure.xlsx":
        print("Enclosure chosen in new")
        Range = "=$DD$2:$DD$" + str(len(CableValidateList)+1)        
        counter = 2
        if CableValidateList:
            CableValidateList.sort()
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,108).Value = item
                counter += 1
            for i in range(1,31):
                Connection = "_Connection"+str(i)
                if Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown == True:
                    Datasheet.ActiveSheet.Range(Connection).Validation.Delete()
            for i in range(1,31):
                Connection = "_Connection"+str(i)
                Datasheet.ActiveSheet.Range(Connection).Validation.Add(3, 1, 3, Range)
                Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown = True 
#---------------------------------------------------------------------------------------------
    if  ObjectName == "DS_BES_Slipring.xlsx":
        print("Slipring chosen in new")
        Range = "=$EE$2:$EE$" + str(len(CableValidateList)+1)        
        counter = 2
        if CableValidateList:
            CableValidateList.sort()
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,135).Value = item
                counter += 1
            for i in range(1,39):
                Connection = "_Connection"+str(i)
                if Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown == True:
                    Datasheet.ActiveSheet.Range(Connection).Validation.Delete()
            for i in range(1,39):
                Connection = "_Connection"+str(i)
                Datasheet.ActiveSheet.Range(Connection).Validation.Add(3, 1, 3, Range)
                Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown = True                 
    Xcel.Visible = True

def choosetag(event):
    print("CHOOSETAG")
    global Tag_List, chosenTag, addr
    objects = []
    New_from_tag_DS = lbox_tags.curselection()
    lbox_tags.itemconfig(New_from_tag_DS[0], background='green',foreground='white')
    if not New_from_tag_DS:
        tk.messagebox.showwarning(title=None, message="First select a Tag,\nor examine a P&ID.")
        return
    chosenTag = Tag_List[New_from_tag_DS[0]]
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    for key in root:
        objects.append(key)
    connection.close()
    storage.close()
    if chosenTag in objects:
        tk.messagebox.showwarning(title=None, message="Object is already in the Database.")
        return
    else:
        chosenTag = Tag_List[New_from_tag_DS[0]]
        lbl_chosenTag["text"] = chosenTag
        
def bulkimport():
    print("BULKIMPORT")
    global OBJ_List, OBJNumbersDict, addr
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()    
    try:
        TagDict = {}
        path = os.getcwd()+"\Tags Export.xlsx"
        if not os.path.exists(path):
            progress.stop()
            tk.messagebox.showwarning(title=None, message="No Tag Export Excel Workbook available.")
            return
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Import = Xcel.Workbooks.Open(path)
        PandID = Import.ActiveSheet.Cells(2,3).Value
        xlUp = -4162
        LastRow = Import.ActiveSheet.Cells(Import.ActiveSheet.Rows.Count, "A").End(xlUp).Row
        for row in range(2,LastRow+1):
            if Import.ActiveSheet.Cells(row,2).Value != None: # To make sure an object has been selected for the tag
                TagDict[Import.ActiveSheet.Cells(row,1).Value] = Import.ActiveSheet.Cells(row,2).Value
            progress.update()
        Import.Close(SaveChanges=True)        
        Xcel.Application.Quit()            
        print(TagDict)
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)             
        db = DB(storage)
        connection = db.open()
        root = connection.root()        
        for tag in TagDict:
            if not(tag in root):
                if TagDict[tag] == "DS_BES_AI":             Object = DS_BES_AI()
                if TagDict[tag] == "DS_BES_AO":             Object = DS_BES_AO()                  
                if TagDict[tag] == "DS_BES_Antenna":        Object = DS_BES_Antenna()
                if TagDict[tag] == "DS_BES_Beacon":         Object = DS_BES_Beacon()
                if TagDict[tag] == "DS_BES_Cable":          Object = DS_BES_Cable()
                if TagDict[tag] == "DS_BES_Compass":        Object = DS_BES_Compass()                
                if TagDict[tag] == "DS_BES_CV":             Object = DS_BES_CV()
                if TagDict[tag] == "DS_BES_DI":             Object = DS_BES_DI()
                if TagDict[tag] == "DS_BES_DO":             Object = DS_BES_DO()                
                if TagDict[tag] == "DS_BES_Enclosure":      Object = DS_BES_Enclosure()
                if TagDict[tag] == "DS_BES_FD":             Object = DS_BES_FD()
                if TagDict[tag] == "DS_BES_FI":             Object = DS_BES_FI()
                if TagDict[tag] == "DS_BES_Fogdetector":    Object = DS_BES_Fogdetector()
                if TagDict[tag] == "DS_BES_Foghorn":        Object = DS_BES_Foghorn()                
                if TagDict[tag] == "DS_BES_GD":             Object = DS_BES_GD()
                if TagDict[tag] == "DS_BES_Handsw":         Object = DS_BES_Handsw()
                if TagDict[tag] == "DS_BES_LG":             Object = DS_BES_LG()
                if TagDict[tag] == "DS_BES_Limitsw":        Object = DS_BES_Limitsw()
                if TagDict[tag] == "DS_BES_LIT":            Object = DS_BES_LIT()
                if TagDict[tag] == "DS_BES_Loadpin":        Object = DS_BES_Loadpin()   
                if TagDict[tag] == "DS_BES_Oceanograph":    Object = DS_BES_Oceanograph()                   
                if TagDict[tag] == "DS_BES_PG":             Object = DS_BES_PG()
                if TagDict[tag] == "DS_BES_PIG":            Object = DS_BES_PIG()        
                if TagDict[tag] == "DS_BES_PIT":            Object = DS_BES_PIT()
                if TagDict[tag] == "DS_BES_RO":             Object = DS_BES_RO()
                if TagDict[tag] == "DS_BES_SOL_V":          Object = DS_BES_SOL_V()
                if TagDict[tag] == "DS_BES_Solarpanel":     Object = DS_BES_Solarpanel()
                if TagDict[tag] == "DS_BES_Slipring":       Object = DS_BES_Slipring()
                if TagDict[tag] == "DS_BES_Speaker":        Object = DS_BES_Speaker()                  
                if TagDict[tag] == "DS_BES_SV":             Object = DS_BES_SV()
                if TagDict[tag] == "DS_BES_TG":             Object = DS_BES_TG()
                if TagDict[tag] == "DS_BES_TT":             Object = DS_BES_TT()
                if TagDict[tag] == "DS_BES_Weatherstation": Object = DS_BES_Weatherstation()
                if TagDict[tag] == "DS_BES_Windgenerator":  Object = DS_BES_Windgenerator()                
                if TagDict[tag] == "DS_BES_Transformer":    Object = DS_BES_Transformer()
                print(tag,Object)
                Object.TagNumber = tag
                Object.PandID = PandID
                root[tag] = Object
                # update OBJNumbersDict
                OBJNumbersDict[type(root[tag]).__name__] += 1
                transaction.commit()
                progress.update()
        OBJ_List = []
        for key in root:
            OBJ_List.append(key)
        # clear the lbox_objects listbox
        lbox_objects.delete(0, last=tk.END)                
        connection.close()
        storage.close()
        OBJ_List.sort()
        for item in OBJ_List:
            lbox_objects.insert(tk.END, item)
        lbl_objectnumber["text"] = "No of Objects: "+str(len(root))          
        progress.stop()        
        tk.messagebox.showwarning(title=None, message="Bulk Import finished.")        
    except BaseException as e:
        print(e.args)
        
def loopimport():
    print("LOOPIMPORT")
    global addr
    answer = tk.messagebox.askyesno(title=None, message="Is the Connections Export Excel Workbook\n"\
                                    "checked on Source (Connection1) and Destination (Connection2)?")
    if answer == False:
        return
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()    
    try:
        CableDict = {}
        MissingTags = []
        MissingCables = []
        Equipment = []
        path = os.getcwd()+"\Connections Export.xlsx"
        if not os.path.exists(path):
            progress.stop()
            tk.messagebox.showwarning(title=None, message="No Connections Export Excel Workbook available.")
            return
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Import = Xcel.Workbooks.Open(path)
        xlUp = -4162
        LastRow = Import.ActiveSheet.Cells(Import.ActiveSheet.Rows.Count, "A").End(xlUp).Row
        for row in range(2,LastRow+1):
            if Import.ActiveSheet.Cells(row,2).Value != None: # To make sure an object has been selected for the tag
                CableDict[Import.ActiveSheet.Cells(row,1).Value] = [Import.ActiveSheet.Cells(row,2).Value,\
                          Import.ActiveSheet.Cells(row,3).Value]
            progress.update()
        Import.Close(SaveChanges=True)        
        Xcel.Application.Quit()
        print(len(CableDict))
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)            
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        #print(CableDict)
        for cable in CableDict:
            item1 = CableDict[cable][0]
            item2 = CableDict[cable][1]
            # print(cable,item1,item2)
            if not item1 in Equipment:
                Equipment.append(item1)
            if not item2 in Equipment:
                Equipment.append(item2)                
            # first fill in the connections for the cables.
            if (cable in root) and isinstance(root[cable], DS_BES_Cable):
                LocalObject = root[cable]
                if (item1 in root) and not isinstance(item1, DS_BES_Cable):
                    # print(cable,item1)                    
                    setattr(LocalObject,'Connection1',item1)
                    root[cable] = LocalObject
                    transaction.commit()
                if (item2 in root) and not isinstance(item2, DS_BES_Cable):
                    # print(cable,item2)
                    setattr(LocalObject,'Connection2',item2)
                    root[cable] = LocalObject
                    transaction.commit()
                if item1 not in root:
                    MissingTags.append(item1)
                if item2 not in root:
                    MissingTags.append(item2)
            # add cables that are not in the database to MissingCables
            # and its equipment to MissingTags
            if not(cable in root):
                MissingCables.append(cable)
                MissingTags.append(item1)
                MissingTags.append(item2)
        # remove missing cables from the CableDict
        for item in MissingCables:
            del CableDict[item]                
        # get the missing tags out of the Equipment list as they are not in the database
        # print(MissingTags)
        # print(MissingCables)
        for item in MissingTags:
            if item in Equipment:
                Equipment.remove(item)
#        # now attach the cables to the equipment
        for item in Equipment:
            LocalObject = root[item]            
            # first enclosures
            if isinstance(LocalObject, DS_BES_Enclosure) or isinstance(LocalObject, DS_BES_Slipring):
                # TODO if Enclosure already has connections should they be overwritten?
                LocalObject.ConnectionNumber = 0
                for cable in CableDict:
                    if item == CableDict[cable][0] or item == CableDict[cable][1]:
                        setattr(LocalObject,'Connection'+str(LocalObject.ConnectionNumber+1),cable)
                        LocalObject.ConnectionNumber += 1
                root[item] = LocalObject
                transaction.commit()
            condition = isinstance(LocalObject,DS_BES_CV) or isinstance(LocalObject,DS_BES_Enclosure) or\
                        isinstance(LocalObject,DS_BES_LG) or isinstance(LocalObject,DS_BES_PG) or\
                        isinstance(LocalObject,DS_BES_TG) or isinstance(LocalObject,DS_BES_RO) or\
                        isinstance(LocalObject,DS_BES_SV) or isinstance(LocalObject,DS_BES_Cable) or\
                        isinstance(LocalObject,DS_BES_Transformer) or isinstance(LocalObject,DS_BES_Slipring)
            # this is equipment that only has maximum 2 connections
            if not (condition):
                for cable in CableDict:
                    if item == CableDict[cable][0] or item == CableDict[cable][1]:
                        if getattr(LocalObject,'Connection1') == "":
                            setattr(LocalObject,'Connection1',cable)
                        elif getattr(LocalObject,'Connection1') != cable and getattr(LocalObject,'Connection2') == "":
                            setattr(LocalObject,'Connection2',cable)                        
        connection.close()
        storage.close()     
        # print(Equipment)
        # combine MissingTags and MissingCables
        MissingTags += MissingCables                   
        if len(MissingTags) > 0:
            MissingTags = list(set(MissingTags))
            text = "There are a number of tags not in the Database.\nDo you want these available in Tag Export.xlx?"
            answer = tk.messagebox.askokcancel(title=None, message=text)
            if answer == True:
                    path = os.getcwd()+"\Tags Export.xlsx"
                    # Remove exisiting collection and create new one
                    if os.path.exists(path):
                        os.remove(path)
                    time.sleep(2) # To give the os module time to remove the Excel Export File    
                    Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
                    Export = Xcel.Workbooks.Add() 
                    Export.SaveAs(path)
                    # Add the headers
                    Export.ActiveSheet.Cells(1,1).Value = "TAG"
                    Export.ActiveSheet.Cells(1,1).Font.Bold = True
                    Export.ActiveSheet.Cells(1,2).Value = "                  OBJECT                  "
                    Export.ActiveSheet.Cells(1,2).Font.Bold = True
                    Export.ActiveSheet.Cells(1,3).Value = "DIAGRAM"
                    Export.ActiveSheet.Cells(1,3).Font.Bold = True    
                    Export.ActiveSheet.Cells(2,3).Value = lbl_diagram["text"]   
                    row = 2
                    for tag in MissingTags:
                        Export.ActiveSheet.Cells(row,1).Value = tag
                        row += 1
                        progress.update()
                    counter = 50
                    for name, obj in inspect.getmembers(sys.modules[__name__]):
                        if inspect.isclass(obj) and obj.__name__.startswith("DS"):
                            Export.ActiveSheet.Cells(counter,108).Value = obj.__name__
                            counter += 1
                    xlUp = -4162
                    LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "DD").End(xlUp).Row
                    Range = "=$DD$50:$DD$" + str(LastRow)
                #    for i in range(1,100):
                #        Cell = "B"+str(i)
                #        if Export.ActiveSheet.Range(Cell).Validation.InCellDropdown == True:
                #            Export.ActiveSheet.Range(Cell).Validation.Delete()
                    for i in range(1,100):
                        Cell = "B"+str(i)
                        Export.ActiveSheet.Range(Cell).Validation.Add(3, 1, 1, Range)
                        Export.ActiveSheet.Range(Cell).Validation.InCellDropdown = True            
                    Export.ActiveSheet.Range("A:EE").Columns.AutoFit()
                    Export.Close(SaveChanges=True)        
                    Xcel.Application.Quit()
        progress.stop()            
        tk.messagebox.showwarning(title=None, message="Connections Import finished.")
    except BaseException as e:
        print("here",e.args)
        
def showclass(event):
    print("SHOWCLASS")
    global addr
    OBJ = lbox_objects.get(lbox_objects.curselection())
    print(OBJ)
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr) 
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    Obj = type(root[OBJ])
    pid = getattr(root[OBJ], "PandID")
    SerDsc = getattr(root[OBJ], "ServiceDescription")
    print(pid)
    connection.close()
    storage.close()
    x = window.winfo_x()
    y = window.winfo_y()
    top = tk.Toplevel(window)
    lbl_objecttype = tk.Label(master=top,text ="Objecttype",justify=tk.LEFT,relief=tk.GROOVE)
    lbl_objecttype.grid(row=0, column=0, sticky="nsew")
    lbl_objecttyp = tk.Label(master=top,text = Obj.__name__,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_objecttyp.grid(row=0, column=1, sticky="nsew")
    lbl_PandID = tk.Label(master=top,text ="PandID",justify=tk.LEFT,relief=tk.GROOVE)
    lbl_PandID.grid(row=1, column=0, sticky="nsew") 
    lbl_PanID = tk.Label(master=top,text =pid,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_PanID.grid(row=1, column=1, sticky="nsew")
    lbl_ServDesc = tk.Label(master=top,text ="ServiceDescription",justify=tk.LEFT,relief=tk.GROOVE)
    lbl_ServDesc.grid(row=2, column=0, sticky="nsew")   
    lbl_ServDes = tk.Label(master=top,text =SerDsc,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_ServDes.grid(row=2, column=1, sticky="nsew") 
    top.geometry("+%d+%d" % (x + 400, y + 600))       
    top.mainloop() 
    return        

def fillobject(event):
    print("FILLOBJECT")
    global SELECTED, addr
    if SELECTED == False: # Apparently no category was selected
       tk.messagebox.showwarning(title=None, message="First select an object Category.")
       return
   # Transfer function for obtaining prefilled data and transfer it to the selected datasheet
    def transfer(selection):
       print("TRANSFER")
       if len(selection) == 0:
           tk.messagebox.showwarning(title=None, message="First select an item.")
           return  
       if len(lbox_prefills.curselection()) == 0: #Transfer can only be done if the user selected a prefilled datasheet
           tk.messagebox.showwarning(title=None, message="First select an item.")
           return
       Datasource = lbox_prefills.get(lbox_prefills.curselection())       
       print(Datasource)
#       directory = os.getcwd()+"/Databases"
#       filename = "/Lufeng.fs"    
#       storage = FileStorage.FileStorage(directory+filename)
       storage = ClientStorage.ClientStorage(addr) 
       db = DB(storage)
       connection = db.open()
       root = connection.root()
       # check whether the selected object already contains manufacturer data
       # and whether the user wants to overwrite this data with new data
       remove_items = []
       for item in selection:
           if root[item].Manufacturer or root[item].Model:
               text = item+" already contains data.\nDo you want to overwrite it?"
               answer = tk.messagebox.askokcancel(title=None, message=text)
               print(answer)
               if not answer:
                   remove_items.append(item)
       if len(remove_items) > 0:
           for item in remove_items:
               selection.remove(item)
       path = os.getcwd()+"\PF Datasheets\\"+type(root[selection[0]]).__name__
       path = path +'\\'+ Datasource
       Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
       PF_Datasheet = Xcel.Workbooks.Open(path)               
       for item in selection:
           Object1 = root[item] # Object1 to prevent issues with global variable Object
           for key in Object1.InternalFieldsDict:
               if PF_Datasheet.ActiveSheet.Range(Object1.InternalFieldsDict[key]).Value != None:
                   print(key)
                   CellValue = PF_Datasheet.ActiveSheet.Range(Object1.InternalFieldsDict[key]).Value
                   setattr(Object1, key, CellValue)
           # Then assign the inherited Dictionary to the corresponding Object Attributes        
           for key in Object1.FieldsDict:
               if PF_Datasheet.ActiveSheet.Range(Object1.FieldsDict[key]).Value != None:
                   print(key)
                   CellValue = PF_Datasheet.ActiveSheet.Range(Object1.FieldsDict[key]).Value
                   setattr(Object1, key, CellValue)
           Object1.TagNumber = item
           root[item] = Object1
           transaction.commit()
       connection.close()
       storage.close()
       PF_Datasheet.Close(SaveChanges=False)
       tk.messagebox.showinfo(title=None, message="Datatransfer finished")
       return
    # First determine the type and tags of the selected object category
    selection = []
    items = lbox_objects.curselection()
    for item in items:
        op = lbox_objects.get(item)
        selection.append(op)
    print(selection)
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"    
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr) 
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    if len(selection) > 0:
        pf_dir = os.getcwd()+"/PF Datasheets/"+type(root[selection[0]]).__name__
    else:
        connection.close()
        storage.close()        
        tk.messagebox.showwarning(title=None, message="First select an object or objects.")
        return
    connection.close()
    storage.close()
    x = window.winfo_x()
    y = window.winfo_y()
    top = tk.Toplevel(window)
    scrb_prefills = tk.Scrollbar(master=top, orient=tk.VERTICAL)
    scrb_prefills.grid(row=0, column=1,  sticky="nsw")
    
    # Object selection via listbox
    lbox_prefills = tk.Listbox(master=top, selectmode=tk.SINGLE, height=18, width=30, yscrollcommand=scrb_tags)
    #lbox_prefills.bind('<Double-1>',choosetag)
    lbox_prefills.grid(row=0, column=0, sticky="nsew", padx=10)
    scrb_prefills['command'] = lbox_prefills.yview
    
    for item in os.listdir(pf_dir):
        lbox_prefills.insert(tk.END, item)
    
    btn_enter = tk.Button(master=top, text="TRANSFER DATA", command=lambda: transfer(selection))
    btn_enter.grid(row=2, column=0, sticky="nsew", padx=10)
    
    top.geometry("+%d+%d" % (x + 900, y + 200))       
    top.mainloop() 
    SELECTED = False
        
def retain():
    print("RETAIN")
    global Datasheet, Object, RETAIN, Spares_List, addr
    global CableTypeDict
    New_OBJ =  lbox_objects.curselection()
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        # print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and not New_OBJ:
            print("No Object chosen")
            tk.messagebox.showwarning(title=None, message="First select an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
#   first compile CableValidateList2core
    Core2or3TypeList = []      
    for key in CableTypeDict:
        coreconfig = CableTypeDict[key][2]
        condition = coreconfig.find('1Pr')==0 or coreconfig.find('1Tr')==0 or\
        coreconfig.find('1Pr')==0 or coreconfig.find('2Pr')==0 or coreconfig.find('2Tr')==0
        if condition:
            Core2or3TypeList.append(coreconfig)
    # print(Core2or3TypeList)
    storage = ClientStorage.ClientStorage(addr) 
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    Object = root[OBJ_List[New_OBJ[0]]]
#    for key in root:
#        print(key)
    connection.close()
    storage.close()    
    print(type(Object).__name__)
    ObjectName = type(Object).__name__ + ".xlsx"  
    path = os.getcwd()+"\Datasheets\\"
    path = path + ObjectName
    Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    Datasheet = Xcel.Workbooks.Open(path)
#    for key in Object.InternalFieldsDict:
#        print(key, getattr(Object, key))
#    for key in Object.FieldsDict:
#        print(key, getattr(Object, key))
    # Assign Object Attribute to the corresponding Cell
    for key in Object.InternalFieldsDict:
        Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value = getattr(Object, key)
        print(Object.InternalFieldsDict[key])
    for key in Object.FieldsDict:
        Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value = getattr(Object, key)
        print(Object.FieldsDict[key])
#---------------------------------------------------------------------------------------------
# prepare validation lists of cables and of instruments for use in drop-down list   
    InstrumentsCount = 1
    CableCount = 1
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    InstrumentsValidateList = []
    CableValidateList = []
    CableValidateList2core = []     
    for key in root:
        obj = root[key]
        if not isinstance(obj,DS_BES_Cable):
            InstrumentsValidateList.append(key)
            InstrumentsCount += 1
        if isinstance(obj,DS_BES_Cable):
            if obj.CoreConfiguration not in Core2or3TypeList:
                CableValidateList.append(key)
                CableCount += 1
            if obj.CoreConfiguration in Core2or3TypeList:
                CableValidateList2core.append(key)
                CableCount += 1                
    connection.close()
    storage.close()
    #print(CableValidateList2core)
#---------------------------------------------------------------------------------------------        
    if  ObjectName == "DS_BES_Cable.xlsx":
        print("Cable chosen in retain")
        # Compile a ModelValidateList from CableTypeDict
        ModelValidateList = []
        for key in CableTypeDict:
            ModelValidateList.append(key)
        # To provide a pull down list in the datasheet of instruments or enclosures that the cables can connect to
        Range1 = "=$EA$2:$EA$" + str(len(InstrumentsValidateList)+1)
        Range2 = "=$EB$2:$EB$" + str(len(ModelValidateList)+1)
        counter = 2
        InstrumentsValidateList.sort()
        if InstrumentsValidateList:
            for item in InstrumentsValidateList:
                Datasheet.ActiveSheet.Cells(counter,131).Value = item # 131 is column EA
                counter += 1
            # The following commands are apparently necessary to insert an updated drop-down list in Excel
            print(Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.InCellDropdown)
            if Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.Delete()
                time.sleep(0.5)
            if Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.Delete()
                time.sleep(0.5)
            # Reconstruction of the updated drop-down list
            Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.Add(3, 1, 3, Range1)
            Datasheet.ActiveSheet.Range("_CnstrConnection1").Validation.InCellDropdown = True
            Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.Add(3, 1, 3, Range1)
            Datasheet.ActiveSheet.Range("_CnstrConnection2").Validation.InCellDropdown = True
            time.sleep(0.5)
        ModelValidateList.sort()
        counter = 2
        if ModelValidateList:
            for item in ModelValidateList:
                Datasheet.ActiveSheet.Cells(counter,132).Value = item # 131 is column EB
                counter += 1
            # The following commands are apparently necessary to insert an updated drop-down list in Excel
            print(Datasheet.ActiveSheet.Range("_Model").Validation.InCellDropdown)
            if Datasheet.ActiveSheet.Range("_Model").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_Model").Validation.Delete()
                time.sleep(0.5)
            # Reconstruction of the updated drop-down list
            Datasheet.ActiveSheet.Range("_Model").Validation.Add(3, 1, 3, Range2)
            Datasheet.ActiveSheet.Range("_Model").Validation.InCellDropdown = True
        # add the cable type parameters
        Model = Datasheet.ActiveSheet.Range("_Model").Value
        if Model in CableTypeDict:
            Datasheet.ActiveSheet.Range("_Manufactr").Value = CableTypeDict[Model][0]
            Datasheet.ActiveSheet.Range("_CnstrCoreConfiguration").Value = CableTypeDict[Model][2]
            Datasheet.ActiveSheet.Range("_ElecVoltage").Value = CableTypeDict[Model][3]
            Datasheet.ActiveSheet.Range("_CnstrCrossSection").Value = CableTypeDict[Model][4]
            Datasheet.ActiveSheet.Range("_CnstrFlameRetardancy").Value = CableTypeDict[Model][5]
            Datasheet.ActiveSheet.Range("_CnstrOverallDiameter").Value = CableTypeDict[Model][6]
            Datasheet.ActiveSheet.Range("_CnstrCableAssembly").Value = CableTypeDict[Model][7]
            Datasheet.ActiveSheet.Range("_ElecChargingCurrent").Value = CableTypeDict[Model][8]
            Datasheet.ActiveSheet.Range("_ElecResistance").Value = CableTypeDict[Model][9]
            Datasheet.ActiveSheet.Range("_ElecCapacitance").Value = CableTypeDict[Model][10]
            Datasheet.ActiveSheet.Range("_CnstrWeight").Value = CableTypeDict[Model][11]
#---------------------------------------------------------------------------------------------          
    # THe following instruments have either no cable, are a cable or have more than 2 cables 
    condition = ObjectName == "DS_BES_CV.xlsx" or ObjectName == "DS_BES_Enclosure.xlsx" or\
                ObjectName == "DS_BES_LG.xlsx" or ObjectName == "DS_BES_PG.xlsx" or\
                ObjectName == "DS_BES_TG.xlsx" or ObjectName == "DS_BES_RO.xlsxx" or\
                ObjectName == "DS_BES_SV.xlsx" or ObjectName == "DS_BES_Cable.xlsx" or\
                ObjectName == "DS_BES_Transformer.xlsx" or ObjectName == "DS_BES_Slipring.xlsx"
    if  not (condition):
        # To provide a pull down list in the datasheet of cables the instruments can connect to 
        print("Instrument chosen in retain")
        Range = "=$EA$2:$EA$" + str(len(CableValidateList)+1)
        Range2 = "=$DZ$2:$DZ$" + str(len(CableValidateList2core)+1)        
        counter = 2
        CableValidateList2core.sort()
        CableValidateList.sort()
        if CableValidateList:
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,131).Value = item # 131 is column EA
                counter += 1
            if Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_Connection1").Validation.Delete()
            if Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown == True:
                Datasheet.ActiveSheet.Range("_Connection2").Validation.Delete()
            Datasheet.ActiveSheet.Range("_Connection1").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown = True
            Datasheet.ActiveSheet.Range("_Connection2").Validation.Add(3, 1, 3, Range)
            Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown = True
        counter = 2
#        if CableValidateList2core:
#            for item in CableValidateList2core:
#                Datasheet.ActiveSheet.Cells(counter,130).Value = item # 130 is column DZ
#                counter += 1
#            if Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown == True:
#                Datasheet.ActiveSheet.Range("_Connection1").Validation.Delete()
#            if Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown == True:
#                Datasheet.ActiveSheet.Range("_Connection2").Validation.Delete()
#            Datasheet.ActiveSheet.Range("_Connection1").Validation.Add(3, 1, 3, Range2)
#            Datasheet.ActiveSheet.Range("_Connection1").Validation.InCellDropdown = True
#            Datasheet.ActiveSheet.Range("_Connection2").Validation.Add(3, 1, 3, Range2)
#            Datasheet.ActiveSheet.Range("_Connection2").Validation.InCellDropdown = True
#---------------------------------------------------------------------------------------------
    if  ObjectName == "DS_BES_Enclosure.xlsx":
        print("Enclosure chosen in retain")
        Range = "=$EA$2:$EA$" + str(len(CableValidateList)+len(Spares_List)+1)  
        Range2= "=$DZ$2:$DZ$" + str(len(CableValidateList2core)+len(Spares_List)+1) 
#        DZ is 130 column
        counter = 2
        CableValidateList2core.sort()
        CableValidateList.sort()
        if CableValidateList:
            # add spares to the CableValidateList because spares only appear in Enclosures
            CableValidateList += Spares_List
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,131).Value = item # 131 is column EA
                counter += 1
            for i in range(17,31):             
                Connection = "_Connection"+str(i)
                if Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown == True:
                    Datasheet.ActiveSheet.Range(Connection).Validation.Delete()
                Datasheet.ActiveSheet.Range(Connection).Validation.Add(3, 1, 3, Range)
                Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown = True 
        counter = 2
        if CableValidateList2core:
            # add spares to the CableValidateList because spares only appear in Enclosures
            CableValidateList2core += Spares_List
            for item in CableValidateList2core:
                Datasheet.ActiveSheet.Cells(counter,130).Value = item # 130 is column DZ
                counter += 1
            for i in range(1,17):        
                Connection = "_Connection"+str(i)
                if Datasheet.ActiveSheet.Range (Connection).Validation.InCellDropdown == True:
                    Datasheet.ActiveSheet.Range (Connection).Validation.Delete()
                Datasheet.ActiveSheet.Range (Connection).Validation.Add(3, 1, 3, Range2)
                Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown = True
                                    
    if  ObjectName == "DS_BES_Slipring.xlsx":
        print("Slipring chosen in retain")
        Range = "=$EA$2:$EA$" + str(len(CableValidateList)+1)        
        counter = 2
        CableValidateList.sort()
        if CableValidateList:
            for item in CableValidateList:
                Datasheet.ActiveSheet.Cells(counter,131).Value = item
                counter += 1
            for i in range(1,39):
                Connection = "_Connection"+str(i)
                if Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown == True:
                    Datasheet.ActiveSheet.Range(Connection).Validation.Delete()
            for i in range(1,39):
                Connection = "_Connection"+str(i)
                Datasheet.ActiveSheet.Range(Connection).Validation.Add(3, 1, 3, Range)
                Datasheet.ActiveSheet.Range(Connection).Validation.InCellDropdown = True                  
    Xcel.Visible = True
    RETAIN = True
#    print(Object.InternalFieldsDict)         

def commit():
    print("COMMIT")
    global Datasheet, Object, OBJ_List, RETAIN, OBJNumbersDict, addr  
    try:
        print(Datasheet.Name)
    except BaseException as e:
        print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",):
            tk.messagebox.showwarning(title=None, message="First select a Datasheet or an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    # Check whether there is a TagNumber else warn the user and return
    if RETAIN == False:
        answer = tk.messagebox.askyesno(title=None, message="Are you sure you want\nto commit this object?")
        if answer == False: # apparently the user doesn't want to commit the object
            for key in Object.InternalFieldsDict:
                Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value = None
            # Then assign the inherited Dictionary to the corresponding Object Attributes        
            for key in Object.FieldsDict:
                Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value = None
            # Finally clear the DD column (might still contain object names)
            for i in range(1,1000):
                Datasheet.ActiveSheet.Cells(i,130).Value = '' # 130 is column DZ                
                Datasheet.ActiveSheet.Cells(i,131).Value = '' # 131 is column EA
            Object = None
            Datasheet.Close(True)
            Datasheet = None            
            return    
    TagNumber = Datasheet.ActiveSheet.Range(Object.FieldsDict["TagNumber"]).Value
    if TagNumber == None: # no tagnumber filled in
        tk.messagebox.showerror(title=None, message="First enter a Tag Number.") #TODO check for correct TagNumber
        return
    # check whether this object already exists and whether it was called from new() not retain()
    # because in the latter case the object obviously already exists and needs to be committed back
    if TagNumber in OBJ_List and RETAIN == False:
        tk.messagebox.showerror(title=None, message="Tag Number already exists\nChange Tag Number")
        return
    # Assign Cell Value from the Datasheet to the corresponding Object Attribute
    # Object has 2 Dictionaries: one inherited from Instrument and one internal.
    # First assign the internal Dictionary to the corresponding Object Attributes
    for key in Object.InternalFieldsDict:
        print(key)
        CellValue = Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value
        setattr(Object, key, CellValue)
    # Then assign the inherited Dictionary to the corresponding Object Attributes        
    for key in Object.FieldsDict:
        print(key)
        CellValue = Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value
        if key == 'TagNumber': # if the user enters a number as tag, this has to be transferred to a string
            CellValue = str(CellValue)
        setattr(Object, key, CellValue)
    ObjectName = Object.TagNumber
     # update the Object Numbers Dictionary
    if RETAIN == False:
        OBJNumbersDict[type(Object).__name__] += 1
        transaction.get().note('New Object '+str(TagNumber))
    # Commit the Object to the Database
    # open the database
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)   
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    root[ObjectName] = Object
    transaction.get().setUser(os.getlogin())
    transaction.commit()
    # Before closing, clear Datasheet
    # TODO check for Prefilled Datasheets which fields have been filled in already
    for key in Object.InternalFieldsDict:
        Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value = None
    # Then assign the inherited Dictionary to the corresponding Object Attributes        
    for key in Object.FieldsDict:
        Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value = None
    # Finally clear the colums in the datasheets that might still contain object names
    for i in range(1,1000):
        Datasheet.ActiveSheet.Cells(i,130).Value = '' # 131 is column DZ        
        Datasheet.ActiveSheet.Cells(i,131).Value = '' # 131 is column EA
        Datasheet.ActiveSheet.Cells(i,131).Value = '' # 132 is column EB
    Object = None
    # Before closing Database connection and storage, update OBJ_list
    OBJ_List = []
    for key in root:
        OBJ_List.append(key)
    lbl_objectnumber["text"] = "No of Objects: "+str(len(root))        
    connection.close()
    storage.close()
    # clear the lbox_objects listbox
    lbox_objects.delete(0, last=tk.END)
    OBJ_List.sort()
    # update the number of objects
    for item in OBJ_List:
        lbox_objects.insert(tk.END, item)    
    Datasheet.Close(True)
    Datasheet = None
    lbl_objectchoice["text"] = "ALL"
    RETAIN = False # reset retain
    os.system('TASKKILL /F /IM excel.exe')
 
def remove():
    print("REMOVE")
    global Datasheet, Object, OBJ_List, OBJNumbersDict, addr     
    selection = []
    items = lbox_objects.curselection()
    for item in items:
        op = lbox_objects.get(item)
        selection.append(op)
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        # print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and (len(selection) == 0):
            print("No Object chosen")
            tk.messagebox.showwarning(title=None, message="First select an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    answer = tk.messagebox.askokcancel(title=None, message="Are you sure you want\nto remove this(these) object(s)?")
    if answer == False:
        return
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
#    print(root)
#    print(OBJ_List)
    for item in selection:
        # print(type(root[item]).__name__)
        # Update OBJNumbersDict
        OBJNumbersDict[type(root[item]).__name__] -= 1
        del root[item]
        transaction.commit()
    # if the object is a cable it probably is connected to an enclosure(s)
    # this means the cable needs to be removed from the enclosure(s) as well        
    for key in root:
        obj = root[key]
        if isinstance(obj,DS_BES_Enclosure):
            print(key)
            flag = False
            for i in range(1,31):    
                attribute = 'Connection'+str(i)
                # first check if the attribute is in the obj.__dict__ because
                # if the Connection is empty in Excel it isn't put in the obj.__dict__
                if obj.__dict__.get(attribute) is not None: 
                    if obj.__dict__[attribute] in selection:
                        obj.__dict__[attribute] = ""
                        flag = True
            if flag: # only commit the object if it has really changed
                root[key] = obj
                transaction.commit()                
    # repopulate the object listbox  
    lbox_objects.delete(0, tk.END)
    OBJ_List = []
    for key in root:
        OBJ_List.append(key)
    OBJ_List.sort()
    for item in OBJ_List:
        lbox_objects.insert(tk.END, item)
    lbl_objectnumber["text"] = "No of Objects: "+str(len(root))        
    connection.close()
    storage.close()
    
def rename():
    print('RENAME')
    global Datasheet, Object, OBJ_List, addr     
    Rename_OBJ =  lbox_objects.curselection()
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        # print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and not Rename_OBJ:
            print("No Object chosen")
            tk.messagebox.showwarning(title=None, message="First select an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    # ask for the new tag number
    new_tag = simpledialog.askstring(title = "Rename Object", prompt = "Enter New Tag Name")
    print(new_tag)
    # no tag name filled in handling
    if new_tag == '':
        tk.messagebox.showwarning(title=None, message="Empty Tagnames not allowed")
        return
    if new_tag in OBJ_List:
        tk.messagebox.showwarning(title=None, message="Object with such a name already exists")
        return       
    # Cancel button pushed
    if new_tag == None:
        print("Renaming Canceled")
        return
    print(Rename_OBJ[0])
    print(OBJ_List[Rename_OBJ[0]])
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
#    print(root)
#    print(OBJ_List)
#    print(OBJ_List[Rename_OBJ[0]])
    root[new_tag] = root[OBJ_List[Rename_OBJ[0]]]
    del root[OBJ_List[Rename_OBJ[0]]]
    Object = root[new_tag]
    Object.TagNumber = new_tag
    transaction.commit()
    # if the object is a cable it probably is connected to an enclosure(s)
    # this means the cable needs to be renamed in the enclosure(s) as well     
    for key in root:
        obj = root[key]
        if isinstance(obj,DS_BES_Enclosure):
            flag = False
            for i in range(1,31):             
                attribute = 'Connection'+str(i)
                #print(obj.__dict__[attribute])
                if obj.__dict__.get(attribute) is not None:
                    if obj.__dict__[attribute] == OBJ_List[Rename_OBJ[0]]:
                        obj.__dict__[attribute] = new_tag
                        flag = True
            if flag:
                root[key] = obj
                transaction.commit()            
    # repopulate the object listbox
    lbox_objects.delete(0, tk.END)
    OBJ_List = []
    for key in root:
        OBJ_List.append(key)
    OBJ_List.sort()
    for item in OBJ_List:
        lbox_objects.insert(tk.END, item)    
    connection.close()
    storage.close()
    
def copyobject():
    print('COPYOBJECT')
    global Datasheet, Object, OBJ_List, addr     
    Copy_OBJ =  lbox_objects.curselection()
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        # print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and not Copy_OBJ:
            print("No Object chosen")
            tk.messagebox.showwarning(title=None, message="First select an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    # ask for the new tag number
    new_tag = simpledialog.askstring(title = "Rename Object", prompt = "Enter Tag Name for Copy")
    print(new_tag)
    # no tag name filled in handling
    if new_tag == '':
        tk.messagebox.showwarning(title=None, message="Empty Tagnames not allowed")
        return
    if new_tag in OBJ_List:
        tk.messagebox.showwarning(title=None, message="Object with such a name already exists")
        return       
    # Cancel button pushed
    if new_tag == None:
        print("Copying Canceled")
        return
    print(Copy_OBJ[0])
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr)
    db = DB(storage)
    connection = db.open()
    root = connection.root()
#    print(root)
#    print(OBJ_List)
#    print(OBJ_List[Rename_OBJ[0]])
    root[new_tag] = copy.deepcopy(root[OBJ_List[Copy_OBJ[0]]])
    # deepcopy function from copy module is used to create a complete independent copy
    Object = root[new_tag]
    Object.TagNumber = new_tag
    transaction.commit()
    # repopulate the object listbox
    lbox_objects.delete(0, tk.END)
    OBJ_List = []
    for key in root:
        OBJ_List.append(key)
    OBJ_List.sort()
    for item in OBJ_List:
        lbox_objects.insert(tk.END, item)    
    connection.close()
    storage.close()
    
def reclassify():
    print("RECLASSIFY")
    global Datasheet, Object, OBJ_List, addr     
    Reclassify_OBJ =  lbox_objects.curselection()
    try:
        print(Datasheet.Name)
        tk.messagebox.showwarning(title=None, message="First commit the Instrument to the Database.")
        return
    except BaseException as e:
        # print(e.args)
        if e.args == ("'NoneType' object has no attribute 'Name'",) and not Reclassify_OBJ:
            print("No Object chosen")
            tk.messagebox.showwarning(title=None, message="First select an Object.")
            return
        if e.args == (-2147023174, 'The RPC server is unavailable.', None, None):
            print("Datasheet closed from Excel")
            tk.messagebox.showerror(title=None, message="Next time use the COMMIT Button to close the Datasheet.\nStart over.")
            Datasheet = None
            Object = None
            return
    Reclassify_OBJ = lbox_objects.get(Reclassify_OBJ)
    print(Reclassify_OBJ)
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
#    storage = FileStorage.FileStorage(directory+filename)
    storage = ClientStorage.ClientStorage(addr) 
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    Obj = type(root[Reclassify_OBJ])
    connection.close()
    storage.close()
    # function that swops the class
    def swopclass(Reclassify_OBJ):
       if len(lbox_prefills.curselection()) == 0: #Transfer can only be done if the user selected a class
           tk.messagebox.showwarning(title=None, message="First select a class.")
           return
       Newclass = lbox_prefills.get(lbox_prefills.curselection())          
#       directory = os.getcwd()+"/Databases"
#       filename = "/Lufeng.fs"
#       storage = FileStorage.FileStorage(directory+filename)
       storage = ClientStorage.ClientStorage(addr) 
       db = DB(storage)
       connection = db.open()
       root = connection.root()
       Obj = type(root[Reclassify_OBJ])
       if Obj.__name__ == Newclass:
           tk.messagebox.showwarning(title=None, message="Object is already of type: "+Newclass)
           connection.close()
           storage.close()
           return
       del root[Reclassify_OBJ]
       transaction.commit()
       if Newclass == "DS_BES_AI":             Object = DS_BES_AI()
       if Newclass == "DS_BES_AO":             Object = DS_BES_AO()       
       if Newclass == "DS_BES_Antenna":        Object = DS_BES_Antenna()
       if Newclass == "DS_BES_Beacon":         Object = DS_BES_Beacon()
       if Newclass == "DS_BES_Cable":          Object = DS_BES_Cable()
       if Newclass == "DS_BES_Compass":        Object = DS_BES_Compass()                
       if Newclass == "DS_BES_CV":             Object = DS_BES_CV()
       if Newclass == "DS_BES_DI":             Object = DS_BES_DI()
       if Newclass == "DS_BES_DO":             Object = DS_BES_DO()       
       if Newclass == "DS_BES_Enclosure":      Object = DS_BES_Enclosure()
       if Newclass == "DS_BES_FD":             Object = DS_BES_FD()
       if Newclass == "DS_BES_FI":             Object = DS_BES_FI()
       if Newclass == "DS_BES_Fogdetector":    Object = DS_BES_Fogdetector()
       if Newclass == "DS_BES_Foghorn":        Object = DS_BES_Foghorn()                
       if Newclass == "DS_BES_GD":             Object = DS_BES_GD()
       if Newclass == "DS_BES_Handsw":         Object = DS_BES_Handsw()
       if Newclass == "DS_BES_LG":             Object = DS_BES_LG()
       if Newclass == "DS_BES_Limitsw":        Object = DS_BES_Limitsw()
       if Newclass == "DS_BES_LIT":            Object = DS_BES_LIT()
       if Newclass == "DS_BES_Loadpin":        Object = DS_BES_Loadpin()
       if Newclass == "DS_BES_Oceanograph":    Object = DS_BES_Oceanograph()                
       if Newclass == "DS_BES_PG":             Object = DS_BES_PG()
       if Newclass == "DS_BES_PIG":            Object = DS_BES_PIG()        
       if Newclass == "DS_BES_PIT":            Object = DS_BES_PIT()
       if Newclass == "DS_BES_RO":             Object = DS_BES_RO()
       if Newclass == "DS_BES_SOL_V":          Object = DS_BES_SOL_V()
       if Newclass == "DS_BES_Solarpanel":     Object = DS_BES_Solarpanel()
       if Newclass == "DS_BES_Slipring":       Object = DS_BES_Slipring()
       if Newclass == "DS_BES_Speaker":        Object = DS_BES_Speaker()                  
       if Newclass == "DS_BES_SV":             Object = DS_BES_SV()
       if Newclass == "DS_BES_TG":             Object = DS_BES_TG()
       if Newclass == "DS_BES_TT":             Object = DS_BES_TT()
       if Newclass == "DS_BES_Weatherstation": Object = DS_BES_Weatherstation()
       if Newclass == "DS_BES_Windgenerator":  Object = DS_BES_Windgenerator()                
       if Newclass == "DS_BES_Transformer":    Object = DS_BES_Transformer()       
       Object.TagNumber = Reclassify_OBJ
       root[Reclassify_OBJ] = Object
       transaction.commit()
       connection.close()
       storage.close()
       tk.messagebox.showinfo(title=None, message="Reclassification finished")
       top.destroy()
       top.update()
       return
   
    x = window.winfo_x()
    y = window.winfo_y()
    top = tk.Toplevel(window)
    scrb_prefills = tk.Scrollbar(master=top, orient=tk.VERTICAL)
    scrb_prefills.grid(row=1, column=1, sticky="nsw")
    
    txt = "Current Class:\n"+Obj.__name__+"\n"+"Select new Class:"
    lbl_classes = tk.Label(master=top,text=txt,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_classes.grid(row=0, column=0, sticky="nsew", padx=10)
    
    # Object selection via listbox
    lbox_prefills = tk.Listbox(master=top, selectmode=tk.SINGLE, height=18, width=30, yscrollcommand=scrb_tags)
    #lbox_prefills.bind('<Double-1>',choosetag)
    lbox_prefills.grid(row=1, column=0, sticky="nsew", padx=10)
    scrb_prefills['command'] = lbox_prefills.yview
    
    for name, obj in inspect.getmembers(sys.modules[__name__]):
        if inspect.isclass(obj) and obj.__name__.startswith("DS"):
            lbox_prefills.insert(tk.END, obj.__name__)
    
    btn_enter = tk.Button(master=top, text="SWOP CLASS", command=lambda: swopclass(Reclassify_OBJ))
    btn_enter.grid(row=2, column=0, columnspan =1, sticky="nsew", padx=10)
    
    top.geometry("+%d+%d" % (x + 900, y + 200))       
    top.mainloop()
      
def index():
    print("INDEX")
    global addr
    btn_index.config(state=tk.DISABLED)
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()    
    try:
        ATTR_List = ['TagNumber','Revision','TagStatus','ProjectScope','ProjectStatus','SystemNumber','SkidName','InstrumentFntn','SequenceNumber','Parallel_Identifier','InstrumentType','ServiceDescription','PandID','LineNumber','CauseEffectDiagram','LocationDrawing','InstrumentAreaDescription','InstrumentAreaCode','Instrument_Skid_Number','FireArea','ControlSystem',
                     'SignalType','IS_NIS','Serial_YN','NE_NDE','ConnectionNode','ConnectionChannel','ConnectionCard','ConnectionCabinetTagNumber','CabinetLocation',
                     'DataSheet','LoopDrawing','LayoutDrawing','Status','DrawingReference','AdditionalInfoRemarks','Section','Manufacturer','Model','CalibratedRangeLRV','CalibratedRangeURV',
                     'ControlAlarmTripsettings_LL','ControlAlarmTripsettings_L','ControlAlarmTripsettings_Ctrl_L', 'ControlAlarmTripsettings_Ctrl_N','ControlAlarmTripsettings_Ctrl_H','ControlAlarmTripsettings_H',
                     'ControlAlarmTripsettings_HH','EngUnit','SensorRange_L','SensorRange_H','SetPointL_N','SetPoint_H','Unit','PO_Number']
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr) 
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        path = os.getcwd()+"\Instrument Master Index.xlsx"
        path2 = os.getcwd()+"\Instrument Index.xlsx"
#         Remove exisiting collection and create new one
        if os.path.exists(path):
            os.remove(path)
        if os.path.exists(path2):
            os.remove(path2)
        time.sleep(2)
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Index = Xcel.Workbooks.Add()
        Ind = Xcel.Workbooks.Add()
        Index.SaveAs(path)
        Ind.SaveAs(path2)
        # Add the headers
        column = 0
        for attribute in ATTR_List:
            progress.update()
            column = column + 1
            Index.ActiveSheet.Cells(1,column).Value = attribute
            Index.ActiveSheet.Cells(1,column).Font.Bold = True
        row = 2   
        for key in root:
            progress.update()
            Object = root[key]
            if not isinstance(Object,DS_BES_Cable) and not isinstance(Object,DS_BES_Enclosure)\
            and not isinstance(Object,DS_BES_Transformer) and not isinstance(Object,DS_BES_Slipring)\
            and not isinstance(Object,DS_BES_Windgenerator) and not isinstance(Object,DS_BES_Transformer)\
            and not isinstance(Object,DS_BES_Solarpanel):
                Index.ActiveSheet.Range("G1:G500").NumberFormat = "@"
                column = 1
                T_list = Object.TagNumber.split("-")
    #                print(T_list)
                if len(T_list) == 4:
                    SN = T_list[3]
                    S_list = SN[:4]
                    SK = T_list[1]
                    DNM = T_list[2]
                    SKDNM =  '-'.join([SK,DNM])
                    IFN = SN[:3]
                    SQN = SN[3:]
#                    print (SKDNM)
#                    print (IF)
#                    print(ZN)
                    for attribute in ATTR_List:
        #                    print(attribute,column)
        #                   for column in range(2,5):
                        if attribute == "SystemNumber":
        #                        print(column,T_list[0])
                            Index.ActiveSheet.Cells(row,column).Value = T_list[0]
                        elif attribute == "SkidName" :
        #                        print(column,T_list[1])
                            Index.ActiveSheet.Cells(row,column).Value = SKDNM
#                            print(T_list[1])
                        elif attribute == "InstrumentFntn":
                            IF = SN[:2]
                            if SN.startswith("H") or SN.startswith("S"):
                                    Index.ActiveSheet.Cells(row,column).Value = IF
                            elif SN.startswith("P"):
                                if len(SN) == 4 :
                                    Index.ActiveSheet.Cells(row,column).Value = IF
                                else:
                                    Index.ActiveSheet.Cells(row,column).Value = IFN
                            else:
                                Index.ActiveSheet.Cells(row,column).Value = IFN
                        elif attribute == "SequenceNumber":
                            ZN = S_list[2:]
                            if SN.startswith("H") or SN.startswith("S"):
                                Index.ActiveSheet.Cells(row,column).Value = ZN
                                print(ZN)
                            elif SN.startswith("P"):
                                if len(SN) == 4:
                                    QN = SN[2:]
                                    Index.ActiveSheet.Cells(row,column).Value = QN
                                else:
                                    Index.ActiveSheet.Cells(row,column).Value = SQN
        #                        print(column,T_list[2])
                            else:
                                Index.ActiveSheet.Cells(row,column).Value = SQN
                        elif attribute == "Parallel_Identifier" :
                            if SN.startswith("H"):
                                PI = SN[4:]
                                Index.ActiveSheet.Cells(row,column).Value = PI 
#                                    print(S_list)
        #                                print(letter)
                            else:
                                Index.ActiveSheet.Cells(row,column).Value = ''
                        else:
                            Index.ActiveSheet.Cells(row,column).Value = getattr(Object, attribute)
                        column = column+1 
                else:
                    for attribute in ATTR_List :
                        if attribute == "SystemNumber":
        #                        print(column,T_list[0])
                            Index.ActiveSheet.Cells(row,column).Value = T_list[0]
                        elif attribute == "SkidName":
        #                        print(column,T_list[1])
                            Index.ActiveSheet.Cells(row,column).Value = ''
                        elif attribute == "InstrumentFntn":
                            Index.ActiveSheet.Cells(row,column).Value = T_list[1]
                        elif attribute == "SequenceNumber":
                            SN = T_list[2]
                            S_list = SN[:4]
                            if SN.startswith("S"):
                                Index.ActiveSheet.Cells(row,column).Value = SN                                
                            else:
                                Index.ActiveSheet.Cells(row,column).Value = S_list 
                        elif attribute == "Parallel_Identifier":
                            SN = T_list[2]
                            if len(SN) > 4 and not SN.startswith("S"):
                                PI = SN[4:]
                                Index.ActiveSheet.Cells(row,column).Value = PI
                                
                            elif len(SN) > 4 and SN.startswith("S"):
                                Index.ActiveSheet.Cells(row,column).Value = ''       
                            else:
                                Index.ActiveSheet.Cells(row,column).Value = '' 
                        else:
                            Index.ActiveSheet.Cells(row,column).Value = getattr(Object, attribute)
                        column = column+1
                row = row + 1
        Index.ActiveSheet.Range("A:CZ").Columns.AutoFit()
        connection.close()
        storage.close()
        Index.Close(SaveChanges=True)
        Xcel.Application.Quit()
        os.system('TASKKILL /F /IM excel.exe')
        tk.messagebox.showwarning(title=None, message="Instrument Master Index prepared.")
        progress.stop()
    except BaseException as e: 
        print(e.args)
        connection.close()
        storage.close()
        os.system('TASKKILL /F /IM excel.exe')
        progress.stop()
        btn_index.config(state=tk.NORMAL)
#create another instrument index with different attributes
    try:
        ATR_list = ['Revision','TagStatus','TagNumber','Module','Skid','TypeCode','Seq_No','Suffix','InstrumentType','ServiceDescription','PandID','ControlSystem','SignalType','IS_NIS','Section','InstrumentAreaDescription',
                    'Instrument_Skid_Number','DataSheet','LoopDrawing','LayoutDrawing','CauseEffectDiagram','Manufacturer','Model','PO_Number','AdditionalInfoRemarks']
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr) 
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        path2 = os.getcwd()+"\Instrument Index.xlsx"
#         Remove exisiting collection and create new one
        if os.path.exists(path2):
            os.remove(path2)
        time.sleep(2)
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Ind = Xcel.Workbooks.Add()
        Ind.SaveAs(path2)
        # Add the headers
        column = 0            
        for attribute in ATR_list:
            column = column + 1
            Ind.ActiveSheet.Cells(1,column).Value = attribute
            Ind.ActiveSheet.Cells(1,column).Font.Bold = True
        row = 2   
        for key in root:
            Object = root[key]
            if not isinstance(Object,DS_BES_Cable) and not isinstance(Object,DS_BES_Enclosure)\
            and not isinstance(Object,DS_BES_Transformer) and not isinstance(Object,DS_BES_Slipring):
                Ind.ActiveSheet.Range("G1:G500").NumberFormat = "@"
                column = 1
                T_list = Object.TagNumber.split("-")
    #                print(T_list)
                if len(T_list) == 4:
                    SN = T_list[3]
                    S_list = SN[:4]
                    SK = T_list[1]
                    DNM = T_list[2]
                    SKDNM =  '-'.join([SK,DNM])
                    IFN = SN[:3]
                    SQN = SN[3:]
#                    print (SKDNM)
#                    print (IF)
#                    print(ZN)
                    for attribute in ATR_list:
        #                    print(attribute,column)
        #                   for column in range(2,5):
                        if attribute == "Module":
        #                        print(column,T_list[0])
                            Ind.ActiveSheet.Cells(row,column).Value = T_list[0]
                        elif attribute == "Skid":
        #                        print(column,T_list[1])
                            Ind.ActiveSheet.Cells(row,column).Value = SKDNM
#                            print(T_list[1])
                        elif attribute == "TypeCode":
                            IF = SN[:2]
                            if SN.startswith("H") or SN.startswith("S"):
                                    Ind.ActiveSheet.Cells(row,column).Value = IF
                            elif SN.startswith("P"):
                                if len(SN) == 4 :
                                    Ind.ActiveSheet.Cells(row,column).Value = IF
                                else:
                                    Ind.ActiveSheet.Cells(row,column).Value = IFN
                            else:
                                Ind.ActiveSheet.Cells(row,column).Value = IFN
                        elif attribute == "Seq_No":
                            ZN = S_list[2:]
                            if SN.startswith("H") or SN.startswith("S"):
                                Ind.ActiveSheet.Cells(row,column).Value = ZN
                                print(ZN)
                            elif SN.startswith("P"):
                                if len(SN) == 4:
                                    QN = SN[2:]
                                    Ind.ActiveSheet.Cells(row,column).Value = QN
                                else:
                                    Ind.ActiveSheet.Cells(row,column).Value = SQN
        #                        print(column,T_list[2])
                            else:
                                Ind.ActiveSheet.Cells(row,column).Value = SQN
                        elif attribute == "Suffix":
                            if SN.startswith("H"):
                                PI = SN[4:]
                                Ind.ActiveSheet.Cells(row,column).Value = PI 
#                                    print(S_list)
        #                                print(letter)
                            else:
                                Ind.ActiveSheet.Cells(row,column).Value = ''
                        else:
                            Ind.ActiveSheet.Cells(row,column).Value = getattr(Object, attribute)
                        column = column+1 
                else:
                    for attribute in ATR_list:
                        if attribute == "Module":
        #                        print(column,T_list[0])
                            Ind.ActiveSheet.Cells(row,column).Value = T_list[0]
                        elif attribute == "Skid":
        #                        print(column,T_list[1])
                            Ind.ActiveSheet.Cells(row,column).Value = ''
                        elif attribute == "TypeCode":
                            Ind.ActiveSheet.Cells(row,column).Value = T_list[1]
                        elif attribute == "Seq_No":
                            SN = T_list[2]
                            S_list = SN[:4]
                            if SN.startswith("S"):
                                Ind.ActiveSheet.Cells(row,column).Value = SN
                            else:
                                Ind.ActiveSheet.Cells(row,column).Value = S_list 
                        elif attribute == "Suffix":
                            SN = T_list[2]
                            if len(SN) > 4 and not SN.startswith("S"):
                                PI = SN[4:]
                                Ind.ActiveSheet.Cells(row,column).Value = PI
                            elif len(SN) > 4 and SN.startswith("S"):
                                Ind.ActiveSheet.Cells(row,column).Value = ''
                            else:
                                Ind.ActiveSheet.Cells(row,column).Value = ''
                        else:
                            Ind.ActiveSheet.Cells(row,column).Value = getattr(Object, attribute)
                        column = column+1
                row = row + 1
#                print(T_list)
#                print(ATTR_List[1:4])
#                print(ATTR_List)
        Ind.ActiveSheet.Range("A:CZ").Columns.AutoFit()
        connection.close()
        storage.close()
        Ind.Close(SaveChanges=True)
        Xcel.Application.Quit()
        os.system('TASKKILL /F /IM excel.exe')
        tk.messagebox.showwarning(title=None, message="Instrument Index prepared.")
        btn_index.config(state=tk.NORMAL)
    except BaseException as e: 
        print(e.args)
        connection.close()
        storage.close()
        os.system('TASKKILL /F /IM excel.exe')
        btn_index.config(state=tk.NORMAL)
        
def startindeximport():
    print("STARTINDEXIMPORTATION")
    threading.Thread(target=Indeximportation).start()
        
def Indeximportation():
    print("INDEXIMPORTATION")
    global addr    
    ATTR_List = ['Revision','TagStatus','ProjectScope','ProjectStatus','SystemNumber','SkidName','InstrumentFntn','SequenceNumber','Parallel_Identifier','InstrumentType','ServiceDescription','PandID','LineNumber','CauseEffectDiagram','LocationDrawing','InstrumentAreaDescription','InstrumentAreaCode','Instrument_Skid_Number','FireArea','ControlSystem',
                     'SignalType','IS_NIS','Serial_YN','NE_NDE','ConnectionNode','ConnectionChannel','ConnectionCard','ConnectionCabinetTagNumber','CabinetLocation',
                     'DataSheet','LoopDrawing','LayoutDrawing','Status','DrawingReference','AdditionalInfoRemarks','Section','Manufacturer','Model','CalibratedRangeLRV','CalibratedRangeURV',
                     'ControlAlarmTripsettings_LL','ControlAlarmTripsettings_L','ControlAlarmTripsettings_Ctrl_L', 'ControlAlarmTripsettings_Ctrl_N','ControlAlarmTripsettings_Ctrl_H','ControlAlarmTripsettings_H',
                     'ControlAlarmTripsettings_HH','EngUnit','SensorRange_L','SensorRange_H','SetPointL_N','SetPoint_H','Unit','PO_Number']
    print("INDEX IMPORT")
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()
    pythoncom.CoInitialize()
    try:
        btn_Inimport.config(state=tk.DISABLED)
        Value_Dict = {}
        TAG_List = []
        path = os.getcwd()+"\Instrument Master Index.xlsx"
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Inimport = Xcel.Workbooks.Open(path)       
        xlToLeft = -4159
        xlUp = -4162
        LastCol = Inimport.ActiveSheet.Cells(1, Inimport.ActiveSheet.Columns.Count).End(xlToLeft).Column
        LastRow = Inimport.ActiveSheet.Cells(Inimport.ActiveSheet.Rows.Count, "A").End(xlUp).Row
        print(LastRow,LastCol)
        for row in range(2,LastRow+1):
            TAG_List.append(Inimport.ActiveSheet.Cells(row,1).Value)
            progress.update()
        row = 2
        for Tag in TAG_List:
#            print(Tag)
#        while row< LastRow+1:
            Temp_list = []
            for column in range(2,LastCol+1):
                Temp_list.append(Inimport.ActiveSheet.Cells(row,column).Value)
                Value_Dict[Tag] = Temp_list
                progress.update()
            row+=1
        Inimport.Close(SaveChanges=True)
        Xcel.Application.Quit()
#        print(TAG_List)             
#        print(Value_Dict)
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for key in Value_Dict:
            LocalObject = root[key]
            for attribute in ATTR_List:
                setattr(LocalObject, attribute, Value_Dict[key][ATTR_List.index(attribute)])
                progress.update()
            root[key] = LocalObject
            transaction.commit()   
        connection.close()
        storage.close()  
        progress.stop()
        pythoncom.CoUninitialize()
        tk.messagebox.showwarning(title=None, message="Instrument Master Index Import finished.")
        btn_Inimport.config(state=tk.NORMAL)
    except BaseException as e:
        print(e.args)
        connection.close()
        storage.close()
        pythoncom.CoUninitialize()
        btn_Inimport.config(state=tk.NORMAL)
        
def terminationexport():
    print("TERMINATIONEXPORT")
    global Spares_List, Spares_Terms
    global CableTypeDict
    global addr
    CableDict = {}
    EquipmentDict = {}
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
            CableDict[key][153] = Start of Color1..50
            CableDict[key][203] Start of RSCR1..50
            CableDict[key][253] = Start of RTermination1..50
            CableDict[key][303] Start of RSignal1..50
            CableDict[key][353] = Start of CoreNumber1..50
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
            # Next compile EquipmentDict
            # This contains the Equipment with for each piece of equipment all cables attached to it.
            # Format EquipmentTag: [ Cable 1, Cable2,....,SPARE,SPARE,..]
            condition = isinstance(obj,DS_BES_AI) or isinstance(obj,DS_BES_AO) or\
                        isinstance(obj,DS_BES_Cable) or isinstance(obj,DS_BES_DI) or\
                        isinstance(obj,DS_BES_DO) or isinstance(obj,DS_BES_LG) or\
                        isinstance(obj,DS_BES_PG) or isinstance(obj,DS_BES_RO) or\
                        isinstance(obj,DS_BES_SV) or isinstance(obj,DS_BES_TG)
            if not condition:
                #print(obj.TagNumber)
                #print(obj.__dict__)
                EquipmentDict[key] = []
                #print(key)
                for i in range(1,31):
                    attribute = 'Connection'+str(i)
                    if attribute in obj.__dict__:
                        #print(attribute,obj.__dict__[attribute])
                        if obj.__dict__[attribute] != "" or obj.__dict__[attribute] != None:
                            EquipmentDict[key].append(obj.__dict__[attribute])
                # Clean EquipmentDict[key] from empty strings
                EquipmentDict[key] = [i for i in EquipmentDict[key] if i]
#        #print( "CableDict Length",len(CableDict[list(CableDict.keys())[0]]))
#        #print(CableDict[list(CableDict.keys())[0]])
        connection.close()
        storage.close()
        #print('EquipmentDict ', EquipmentDict)
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
        for key in EquipmentDict:
            progress.update()
            startrow = row
            Terminations.ActiveSheet.Cells(row,9).Value = key
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"M"+str(row)).Merge() 
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"M"+str(row)).Interior.ColorIndex = 15 
            Terminations.ActiveSheet.Range("I"+str(row)+":"+"M"+str(row)).Font.Bold = True
            row += 1
            for item in EquipmentDict[key]: # item is a Cable Tag or None
                progress.update()
                if item != None:
                    startrow2 = row
                    cable = item
                    # First define the header for the cable: "Signal","Term#","Core#","Configuration","Core#","Term#","Signal"
                    Terminations.ActiveSheet.Cells(row,9).Value = Header2[0]
                    Terminations.ActiveSheet.Cells(row,10).Value = Header2[1]
                    Terminations.ActiveSheet.Cells(row,11).Value = Header2[2]
                    Terminations.ActiveSheet.Cells(row,12).Value = cable # cable tag
                    Terminations.ActiveSheet.Cells(row,12).Interior.ColorIndex = 4
#                    Terminations.ActiveSheet.Cells(row,13).Value = Header2[4]
#                    Terminations.ActiveSheet.Cells(row,14).Value = Header2[5]
#                    Terminations.ActiveSheet.Cells(row,15).Value = Header2[6]
                    condition1 = CableDict[item][2] != key # Enclosure is on left side of the cable
                    condition2 = CableDict[item][2] == key # Enclosure is on right side of the cable
                    if condition1:
                        Terminations.ActiveSheet.Cells(row,13).Value = CableDict[cable][2]
                    if condition2:
                        Terminations.ActiveSheet.Cells(row,13).Value = CableDict[cable][0]
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
#                                Color = Name + "Color" + str(i)                       
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
#                                Terminations.ActiveSheet.Cells(row,13).Value = CableDict[item][i+352] # CoreNumber starts at index 353
#                                if condition1:
#                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
#                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+252]  # RTermination starts on index 253
#                                    Terminations.ActiveSheet.Cells(row,14).Name = RTermination
#                                if condition2:
#                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
#                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+52]  # LTermination starts on index 53
#                                    Terminations.ActiveSheet.Cells(row,14).Name = LTermination
#                                if condition1:
#                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+302]  # RSignal starts on index 303
#                                    Terminations.ActiveSheet.Cells(row,15).Name = RSignal  
#                                if condition2:
#                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+2]  # LSignal starts on index 3
#                                    Terminations.ActiveSheet.Cells(row,15).Name = LSignal                                    
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
#                                    Terminations.ActiveSheet.Cells(row,13).Value = "SCR"+str(screennumber)
#                                    if condition1:
#                                        Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
#                                        Terminations.ActiveSheet.Cells(row,14).Name = RSCR
#                                    if condition2:
#                                        Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+102]  # LSCR starts on index 103
#                                        Terminations.ActiveSheet.Cells(row,14).Name = LSCR                                        
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
#                                Color = Name + "Color" + str(i)                       
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
#                                Terminations.ActiveSheet.Cells(row,13).Value = CableDict[item][i+352] #  CoreNumber starts at index 353
#                                if condition1:
#                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
#                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+252]  # RTermination starts on index 253
#                                    Terminations.ActiveSheet.Cells(row,14).Name = RTermination
#                                if condition2:
#                                    Terminations.ActiveSheet.Range("N"+str(row)).NumberFormat = "@"
#                                    Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][i+52]  # LTermination starts on index 53
#                                    Terminations.ActiveSheet.Cells(row,14).Name = LTermination
#                                if condition1:                                    
#                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+302]  # RSignal starts on index 303
#                                    Terminations.ActiveSheet.Cells(row,15).Name = RSignal
#                                if condition2:
#                                    Terminations.ActiveSheet.Cells(row,15).Value = CableDict[item][i+2]  # LSignal starts on index 3
#                                    Terminations.ActiveSheet.Cells(row,15).Name = LSignal                                   
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
#                            Terminations.ActiveSheet.Cells(row,13).Value = screenconfig
#                            if condition1:
#                                Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
#                                Terminations.ActiveSheet.Cells(row,14).Name = RSCR
#                            if condition2:
#                                Terminations.ActiveSheet.Cells(row,14).Value = CableDict[item][screennumber+102]  # LSCR starts on index 103
#                                Terminations.ActiveSheet.Cells(row,14).Name = LSCR                                
                            row +=1
                    if type(CableDict[item][1]) is str:
                        Terminations.ActiveSheet.Cells(row,12).Value = CableDict[item][1]
                        row += 1
                    HandledCables.append(item)
                    Terminations.ActiveSheet.Range("M"+str(startrow2)+":"+"M"+str(row-1)).Merge()
                    Terminations.ActiveSheet.Range("M"+str(startrow2)+":"+"M"+str(row-1)).VerticalAlignment = -4160 # Align top
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(9).LineStyle = 1 # Continous line
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(7).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(10).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(9).Weight = 4 # Thick linestyle
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(7).Weight = 4
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(10).Weight = 4
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(12).LineStyle = 1 # internal cell borders
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(11).LineStyle = 1
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(12).Weight = 2 # internal cell borders Thin
            Terminations.ActiveSheet.Range("I"+str(startrow)+":"+"M"+str(row-1)).Borders(11).Weight = 2
        Terminations.ActiveSheet.Range("A:M").Columns.AutoFit()
        Terminations.ActiveSheet.Range("A:MP").HorizontalAlignment = xlHAlignCenter            
        # Add the macro for automatic cell changing
        xlButtonControl = 0     
        pathToMacro = os.getcwd()+"\Macro.txt"
        with open (pathToMacro, "r") as myfile:
            print('reading macro into string from: ' + str(myfile))
            macro2=myfile.read()
        excelModule = Terminations.VBProject.VBComponents.Add(1)
        excelModule.CodeModule.AddFromString(macro2)
        StartButton = Terminations.ActiveSheet.Shapes.AddFormControl(xlButtonControl,539.25,15.75,43.5,33.75)
        StartButton.OnAction = "Button1_Click"
        Terminations.ActiveSheet.Shapes("Button 1").Left = Terminations.ActiveSheet.Range("H1").Left 
        Terminations.ActiveSheet.Shapes("Button 1").Top = Terminations.ActiveSheet.Range("H1").Top
        Terminations.ActiveSheet.Shapes("Button 1").TextFrame.Characters(1,8).Text = ""
        Terminations.ActiveSheet.Shapes("Button 1").TextFrame.Characters(1,1).Text = "START"
        UpdateButton = Terminations.ActiveSheet.Shapes.AddFormControl(xlButtonControl,539.25,15.75,43.5,33.75)
        UpdateButton.OnAction = "Button2_Click"
        Terminations.ActiveSheet.Shapes("Button 2").Left = Terminations.ActiveSheet.Range("H4").Left 
        Terminations.ActiveSheet.Shapes("Button 2").Top = Terminations.ActiveSheet.Range("H4").Top
        Terminations.ActiveSheet.Shapes("Button 2").TextFrame.Characters(1,8).Text = ""
        Terminations.ActiveSheet.Shapes("Button 2").TextFrame.Characters(1,1).Text = "UPDATE"   
        time.sleep(2)
        Terminations.ActiveSheet.Protect(DrawingObjects=False)
        #Xcel.Application.Run(MacroName)    
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
        
def terminationimport():
    print("TERMINATIONIMPORT")
    global addr
    # first obtain all cables from the database
    LSignalDict = {}
    LTerminationDict = {}
    LSCRDict = {}
    ColorDict = {}
    RSCRDict = {}
    RTerminationDict = {}
    RSignalDict = {}    
#    directory = os.getcwd()+"/Databases"
#    filename = "/Lufeng.fs"
    progress.start()
    progress.update()    
#    storage = FileStorage.FileStorage(directory+filename) 
    storage = ClientStorage.ClientStorage(addr)             
    db = DB(storage)
    connection = db.open()
    root = connection.root()
    for key in root:
        obj = root[key]
        progress.update()
        if isinstance(obj,DS_BES_Cable):
            LSignalDict[key] = []
            LTerminationDict[key] = []
            LSCRDict[key] = []
            ColorDict[key] = []
            RSCRDict[key] = []
            RTerminationDict[key] = []
            RSignalDict[key] = []
    connection.close()
    storage.close()
    try:
        btn_terminationexport.config(state=tk.DISABLED)
        btn_terminationimport.config(state=tk.DISABLED)        
        path = os.getcwd()+"\Terminations.xlsm"
        if not os.path.exists(path):
            progress.stop()
            tk.messagebox.showwarning(title=None, message="No Tag Export Excel Workbook available.")
            return
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Terminations = Xcel.Workbooks.Open(path)
        # Start filling the Dictionaries
        numerator = 1
        for name in Terminations.Names:
            if name.Name.startswith("aa"):
                for cable in LSignalDict:
                    print('\r'+str(round((numerator/len(Terminations.Names))))+'%', end='\r') # indicate percentage progress
                    numerator += 1
                    progress.update()
                    #print("Handling cable: " + cable)
                    cablename = "aa" + cable.replace('/','') # remove forwardslashes from cable names
                    cablename = cablename.replace('-','') # remove dashes from cable names
                    LSignal = cablename + "LSignal"
                    if LSignal in name.Name:
                        if name.Name[-2:].isnumeric():
                            LSignalDict[cable].append(name.Name[-2:]) #add LSignal number (2 digits)
                        else:
                            LSignalDict[cable].append(name.Name[-1:]) #add LSignal number (1# digit)
                        LSignalDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    LTermination = cablename + "LTermination"
                    if LTermination in name.Name:
                        if name.Name[-2:].isnumeric():
                            LTerminationDict[cable].append(name.Name[-2:]) 
                        else:
                            LTerminationDict[cable].append(name.Name[-1:])
                        LTerminationDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    LSCR = cablename + "LSCR"
                    if LSCR in name.Name:
                        if name.Name[-2:].isnumeric():
                            LSCRDict[cable].append(name.Name[-2:])
                        else:
                            LSCRDict[cable].append(name.Name[-1:])
                        LSCRDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    Color = cablename + "Color"
                    if Color in name.Name:
                        if name.Name[-2:].isnumeric():
                            ColorDict[cable].append(name.Name[-2:])
                        else:
                            ColorDict[cable].append(name.Name[-1:])
                        ColorDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    RSCR = cablename + "RSCR"
                    if RSCR in name.Name:
                        if name.Name[-2:].isnumeric():
                            RSCRDict[cable].append(name.Name[-2:])
                        else:
                            RSCRDict[cable].append(name.Name[-1:])
                        RSCRDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    RTermination = cablename + "RTermination"
                    if RTermination in name.Name:
                        if name.Name[-2:].isnumeric():
                            RTerminationDict[cable].append(name.Name[-2:])
                        else:
                            RTerminationDict[cable].append(name.Name[-1:])
                        RTerminationDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break
                    RSignal = cablename + "RSignal"
                    if RSignal in name.Name:
                        if name.Name[-2:].isnumeric():
                            RSignalDict[cable].append(name.Name[-2:])
                        else:
                            RSignalDict[cable].append(name.Name[-1:])
                        RSignalDict[cable].append(Terminations.ActiveSheet.Range(name).Value)
                        break                                                          
        Terminations.Close(SaveChanges=True)
        Xcel.Application.Quit()
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)           
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for cable in LSignalDict:
            #print('\r'+str(round(50+(numerator/len(LSignalDict))*50))+'%', end='\r')
            #numerator += 1            
            progress.update()
            localobject = root[cable]
            LSignal = LSignalDict[cable]
            for i in range(0,len(LSignal),2):
                attribute = "LSignal"+LSignal[i]
                setattr(localobject,attribute,LSignal[i+1])
            LTermination = LTerminationDict[cable]
            for i in range(0,len(LTermination),2):
                attribute = "LTermination"+LTermination[i]
                setattr(localobject,attribute,LTermination[i+1])
            LSCR = LSCRDict[cable]
            for i in range(0,len(LSCR),2):
                attribute = "LSCR"+LSCR[i]
                setattr(localobject,attribute,LSCR[i+1]) 
            Color = ColorDict[cable]
            for i in range(0,len(Color),2):
                attribute = "Color"+Color[i]
                setattr(localobject,attribute,Color[i+1])
            RSCR = RSCRDict[cable]
            for i in range(0,len(RSCR),2):
                attribute = "RSCR"+RSCR[i]
                setattr(localobject,attribute,RSCR[i+1])
            RTermination = RTerminationDict[cable]
            for i in range(0,len(LTermination),2):
                attribute = "RTermination"+RTermination[i]
                setattr(localobject,attribute,RTermination[i+1])
            RSignal = RSignalDict[cable]
            for i in range(0,len(RSignal),2):
                attribute = "RSignal"+RSignal[i]
                setattr(localobject,attribute,RSignal[i+1])
            root[cable] = localobject
            transaction.commit()               
        connection.close()
        storage.close()
        progress.stop()
        tk.messagebox.showwarning(title=None, message="Terminations Import finished.")
        btn_terminationexport.config(state=tk.NORMAL)
        btn_terminationimport.config(state=tk.NORMAL)   
    except BaseException as e:
        print(e.args)
        connection.close()
        storage.close()        
        progress.stop()
        btn_terminationexport.config(state=tk.NORMAL)
        btn_terminationimport.config(state=tk.NORMAL)

def objects():
    print("OBJECTS")
    global OBJ_List, SELECTED, addr
    try:
        print(ObjectChoice.get())
        SELECTED = True
        lbl_objectchoice["text"] = ObjectChoice.get()
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        # repopulate the object listbox
        lbox_objects.delete(0, tk.END)
        OBJ_List = []
        for key in root:
            if ObjectChoice.get() == "ALL":
                OBJ_List.append(key)
                SELECTED = False
            if ObjectChoice.get() == "ANALOG INPUTS" and isinstance(root[key],DS_BES_AI):
                OBJ_List.append(key)
            if ObjectChoice.get() == "ANALOG OUTPUTS" and isinstance(root[key],DS_BES_AO):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "ANTENNAS" and isinstance(root[key],DS_BES_Antenna):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "BEACONS" and isinstance(root[key],DS_BES_Beacon):
                OBJ_List.append(key)
            if ObjectChoice.get() == "CABLES" and isinstance(root[key],DS_BES_Cable):
                OBJ_List.append(key)
            if ObjectChoice.get() == "COMPASSES" and isinstance(root[key],DS_BES_Compass):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "CONTROL VALVES" and isinstance(root[key],DS_BES_CV):
                OBJ_List.append(key)
            if ObjectChoice.get() == "DIGITAL INPUTS" and isinstance(root[key],DS_BES_DI):
                OBJ_List.append(key)
            if ObjectChoice.get() == "DIGITAL OUTPUTS" and isinstance(root[key],DS_BES_DO):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "ENCLOSURES" and isinstance(root[key],DS_BES_Enclosure):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "FIRE DETECTORS" and isinstance(root[key],DS_BES_FD):
                OBJ_List.append(key)
            if ObjectChoice.get() == "FLOWMETERS" and isinstance(root[key],DS_BES_FI):
                OBJ_List.append(key)
            if ObjectChoice.get() == "FOGDETECTORS" and isinstance(root[key],DS_BES_Fogdetector):
                OBJ_List.append(key) 
            if ObjectChoice.get() == "FOGHORNS" and isinstance(root[key],DS_BES_Foghorn):
                OBJ_List.append(key)                  
            if ObjectChoice.get() == "GAS DETECTORS" and isinstance(root[key],DS_BES_GD):
                OBJ_List.append(key)
            if ObjectChoice.get() == "HANDSWITCHES" and isinstance(root[key],DS_BES_Handsw):
                OBJ_List.append(key)
            if ObjectChoice.get() == "LEVEL GAUGES" and isinstance(root[key],DS_BES_LG):
                OBJ_List.append(key)
            if ObjectChoice.get() == "LIMIT SWITCHES" and isinstance(root[key],DS_BES_Limitsw):
                OBJ_List.append(key)
            if ObjectChoice.get() == "LEVEL TRANSMITTERS" and isinstance(root[key],DS_BES_LIT):
                OBJ_List.append(key)
            if ObjectChoice.get() == "LOADPINS" and isinstance(root[key],DS_BES_Loadpin):
                OBJ_List.append(key)
            if ObjectChoice.get() == "OCEANOGRAPHS" and isinstance(root[key],DS_BES_Oceanograph):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "PRESSURE GAUGES" and isinstance(root[key],DS_BES_PG):
                OBJ_List.append(key)  
            if ObjectChoice.get() == "PIG DETECTORS" and isinstance(root[key],DS_BES_PIG):
                OBJ_List.append(key)
            if ObjectChoice.get() == "PRESSURE TRANSMITERS" and isinstance(root[key],DS_BES_PIT):
                OBJ_List.append(key)
            if ObjectChoice.get() == "RESTRICTION ORIFICES" and isinstance(root[key],DS_BES_RO):
                OBJ_List.append(key)
            if ObjectChoice.get() == "SLIPRINGS" and isinstance(root[key],DS_BES_Slipring):
                OBJ_List.append(key)
            if ObjectChoice.get() == "SPEAKERS" and isinstance(root[key],DS_BES_Speaker):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "SOLENOIDS" and isinstance(root[key],DS_BES_SOL_V):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "SOLARPANELS" and isinstance(root[key],DS_BES_Solarpanel):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "SAFETY VALVES" and isinstance(root[key],DS_BES_SV):
                OBJ_List.append(key)
            if ObjectChoice.get() == "TEMPERATURE GAUGES" and isinstance(root[key],DS_BES_TG):
                OBJ_List.append(key)
            if ObjectChoice.get() == "TEMPERATURE TRANSMITTERS" and isinstance(root[key],DS_BES_TT):
                OBJ_List.append(key)
            if ObjectChoice.get() == "WEATHERSTATIONS" and isinstance(root[key],DS_BES_Weatherstation):
                OBJ_List.append(key)
            if ObjectChoice.get() == "WINDGENERATORS" and isinstance(root[key],DS_BES_Windgenerator):
                OBJ_List.append(key)                
            if ObjectChoice.get() == "TRANSFORMERS" and isinstance(root[key],DS_BES_Transformer):
                OBJ_List.append(key)                
        OBJ_List.sort()
        for item in OBJ_List:
            lbox_objects.insert(tk.END, item)    
        connection.close()
        storage.close()
    except BaseException as e: 
        print(e.args)
        connection.close()
        storage.close()        

def startexport():
    threading.Thread(target=exportation).start()        
    
def exportation():
    print("EXPORTATION")
    pythoncom.CoInitialize()
    global addr
    global CableTypeDict
    global OBJ_List, Spares_list
    try:
#        print(OBJ_List)
#        print(ObjectChoice.get())
        btn_export.config(state=tk.DISABLED)
        btn_import.config(state=tk.DISABLED)
        progress['mode'] = 'indeterminate'
        progress.start()
        progress.update()
        if lbl_objectchoice['text'] == "ALL":
            tk.messagebox.showwarning(title=None, message="First Select a Specific Object Type.")
            progress.stop()
            btn_export.config(state=tk.NORMAL)
            btn_import.config(state=tk.NORMAL)
            return
        LocalObject = None
        # here the class is assigned to LocalObject, based on the choice from the ObjectChoice pulldown list.
        if ObjectChoice.get() == "ANALOG INPUTS":
            LocalObject = DS_BES_AI()
        if ObjectChoice.get() == "ANALOG OUTPUTS":
            LocalObject = DS_BES_AO()        
        if ObjectChoice.get() == "ANTENNAS":
            LocalObject = DS_BES_Antenna()        
        if ObjectChoice.get() == "BEACONS":
            LocalObject = DS_BES_Beacon()
        if ObjectChoice.get() == "CABLES":
            LocalObject = DS_BES_Cable()
        if ObjectChoice.get() == "COMPASSES":
            LocalObject = DS_BES_Compass()            
        if ObjectChoice.get() == "CONTROL VALVES":
            LocalObject = DS_BES_CV()
        if ObjectChoice.get() == "DIGITAL INPUTS":
            LocalObject = DS_BES_DI()
        if ObjectChoice.get() == "DIGITAL OUTPUTS":
            LocalObject = DS_BES_DO()            
        if ObjectChoice.get() == "ENCLOSURES":
            LocalObject = DS_BES_Enclosure()            
        if ObjectChoice.get() == "FIRE DETECTORS":
            LocalObject = DS_BES_FD()
        if ObjectChoice.get() == "FLOWMETERS":
            LocalObject = DS_BES_FI()
        if ObjectChoice.get() == "FOGDETECTORS":
            LocalObject = DS_BES_Fogdetector()
        if ObjectChoice.get() == "FOGHORNS":
            LocalObject = DS_BES_Foghorn()            
        if ObjectChoice.get() == "GAS DETECTORS":
            LocalObject = DS_BES_GD()
        if ObjectChoice.get() == "HANDSWITCHES":
            LocalObject = DS_BES_Handsw()
        if ObjectChoice.get() == "LEVEL GAUGES":
            LocalObject = DS_BES_LG()
        if ObjectChoice.get() == "LIMIT SWITCHES":
            LocalObject = DS_BES_Limitsw()
        if ObjectChoice.get() == "LEVEL TRANSMITTERS":
            LocalObject = DS_BES_LIT()
        if ObjectChoice.get() == "LOADPINS":
            LocalObject = DS_BES_Loadpin()  
        if ObjectChoice.get() == "OCEANOGRAPHS":
            LocalObject = DS_BES_Oceanograph()              
        if ObjectChoice.get() == "PRESSURE GAUGES":
            LocalObject = DS_BES_PG()
        if ObjectChoice.get() == "PIG DETECTORS":
            LocalObject = DS_BES_PIG()
        if ObjectChoice.get() == "PRESSURE TRANSMITERS":
            LocalObject = DS_BES_PIT()
        if ObjectChoice.get() == "RESTRICTION ORIFICES":
            LocalObject = DS_BES_RO()
        if ObjectChoice.get() == "SLIPRINGS":
            LocalObject = DS_BES_Slipring()  
        if ObjectChoice.get() == "SPEAKERS":
            LocalObject = DS_BES_Speaker()              
        if ObjectChoice.get() == "SOLENOIDS":
            LocalObject = DS_BES_SOL_V()
        if ObjectChoice.get() == "SOLARPANELS":
            LocalObject = DS_BES_Solarpanel()            
        if ObjectChoice.get() == "SAFETY VALVES":
            LocalObject = DS_BES_SV()
        if ObjectChoice.get() == "TEMPERATURE GAUGES":
            LocalObject = DS_BES_TG()
        if ObjectChoice.get() == "TEMPERATURE TRANSMITTERS":
            LocalObject = DS_BES_TT()
        if ObjectChoice.get() == "WEATHERSTATIONS":
            LocalObject = DS_BES_Weatherstation() 
        if ObjectChoice.get() == "WINDGENERATORS":
            LocalObject = DS_BES_Windgenerator()             
        if ObjectChoice.get() == "TRANSFORMERS":
            LocalObject = DS_BES_Transformer() 
        chosen_object = ObjectChoice.get()
        path = os.getcwd()+"\Object Export.xlsx"
        # Remove exisiting collection and create new one
        if os.path.exists(path):
            os.remove(path)
        time.sleep(2) # To give the os module time to remove the Excel Export File
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Export = Xcel.Workbooks.Add() 
        Export.SaveAs(path)
        Xcel.Visible = False
        # Add the headers in the first row of the spreadsheet
        column = 1
        for key in LocalObject.FieldsDict:  #######
            Export.ActiveSheet.Cells(1,column).Value = key
            Export.ActiveSheet.Cells(1,column).Font.Bold = True
            column = column + 1
            progress.update()
        for key in LocalObject.InternalFieldsDict:
            Export.ActiveSheet.Cells(1,column).Value = key
            Export.ActiveSheet.Cells(1,column).Font.Bold = True
            column = column + 1
            progress.update()
        row = 2
        column = 1
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        # Filling the spreadsheet cells with the attribute values of the chosen object
        storage = ClientStorage.ClientStorage(addr) 
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for item in OBJ_List: # OBJ_List is only filled with those object chosen in function objects
            LocalObject = root[item]
            for key in LocalObject.FieldsDict:
#                if not key == "Date":
                Export.ActiveSheet.Cells(row,column).Value = getattr(LocalObject, key)
                column = column + 1
            for key in LocalObject.InternalFieldsDict:
                Export.ActiveSheet.Cells(row,column).Value = getattr(LocalObject, key)
                column = column + 1
            row = row + 1
            column = 1
            progress.update()
        connection.close()
        storage.close()
#---------------------------------------------------------------------------------------------
# prepare validation lists of cables for use in drop-down list 
        Core2or3TypeList = []      
        for key in CableTypeDict:
            coreconfig = CableTypeDict[key][2]
            condition = coreconfig.find('1Pr')==0 or coreconfig.find('1Tr')==0 or\
            coreconfig.find('1Pr')==0 or coreconfig.find('2Pr')==0 or coreconfig.find('2Tr')==0
            if condition:
                Core2or3TypeList.append(coreconfig)        
        if isinstance(LocalObject, DS_BES_Enclosure):
            CableCount = 1
#            directory = os.getcwd()+"/Databases"
#            filename = "/Lufeng.fs"
#            storage = FileStorage.FileStorage(directory+filename)
            storage = ClientStorage.ClientStorage(addr) 
            db = DB(storage)
            connection = db.open()
            root = connection.root()
            CableValidateList = []
            CableValidateList2core = []
            for key in root:
                obj = root[key]
                if isinstance(obj,DS_BES_Cable):
                    if obj.CoreConfiguration not in Core2or3TypeList:
                        CableValidateList.append(key)
                        CableCount += 1
                    if obj.CoreConfiguration in Core2or3TypeList:
                        CableValidateList2core.append(key)
                        CableCount += 1
            connection.close()
            storage.close()
#---------------------------------------------------------------------------------------------
        # To provide a pull down list in the datasheet of instruments or enclosures that the cables can connect to
            Range = "=$DD$2:$DD$" + str(len(CableValidateList)+len(Spares_List)+1)
            Range2= "=$DC$2:$DC$" + str(len(CableValidateList2core)+len(Spares_List)+1)
            counter = 2
            CableValidateList2core.sort()
            CableValidateList.sort()
            if CableValidateList2core:
                # add spares to the CableValidateList because spares only appear in Enclosures
                CableValidateList2core += Spares_List             
                for item in CableValidateList2core:
                    Export.ActiveSheet.Cells(counter,107).Value = item # 107 is column DC
                    counter += 1
                xlUp = -4162
                LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "A").End(xlUp).Row
                for row in range(2,LastRow+1):
                    for column in range(51,68):
                       # Construction of the updated drop-down list
                       Export.ActiveSheet.Cells(row,column).Validation.Add(3, 1, 3, Range2)
                       Export.ActiveSheet.Cells(row,column).Validation.InCellDropdown = True 
            counter = 2                       
            if CableValidateList:
                # add spares to the CableValidateList because spares only appear in Enclosures
                CableValidateList += Spares_List             
                for item in CableValidateList:
                    Export.ActiveSheet.Cells(counter,108).Value = item # 108 is column DD
                    counter += 1
                xlUp = -4162
                LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "A").End(xlUp).Row
                for row in range(2,LastRow+1):
                    for column in range(68,81):
                       # Construction of the updated drop-down list
                       Export.ActiveSheet.Cells(row,column).Validation.Add(3, 1, 3, Range)
                       Export.ActiveSheet.Cells(row,column).Validation.InCellDropdown = True           
#---------------------------------------------------------------------------------------------
# prepare validation lists of instruments and enclosures for use in drop-down list 
        if isinstance(LocalObject, DS_BES_Cable):
            InstrumentsCount = 1
#            directory = os.getcwd()+"/Databases"
#            filename = "/Lufeng.fs"
#            storage = FileStorage.FileStorage(directory+filename)
            storage = ClientStorage.ClientStorage(addr) 
            db = DB(storage)
            connection = db.open()
            root = connection.root()
            InstrumentsValidateList = []
            for key in root:
                obj = root[key]
                if not isinstance(obj,DS_BES_Cable):
                    InstrumentsValidateList.append(key)
                    InstrumentsCount += 1
                    print(key)
            connection.close()
            storage.close()
            # Prepare a cabetype list
            CableTypeList = []
            for cableType in CableTypeDict:
                CableTypeList.append(cableType)
            # Prepare a screen configuration list
            ScreenList = ["No Screen","Overall Screen","Screen per Pair","Screen per Triad"]
#---------------------------------------------------------------------------------------------
        # To provide a pull down list in the datasheet of instruments or enclosures that the cables can connect to
            Range = "=$DD$2:$DD$" + str(len(InstrumentsValidateList)+1)
            Range2 = "=$DE$2:$DE$" + str(len(CableTypeList)+1)
            Range3 = "=$DF$2:$DF$" + str(len(ScreenList)+1)
            counter = 2
            InstrumentsValidateList.sort()
            if InstrumentsValidateList:
                for item in InstrumentsValidateList:
                    Export.ActiveSheet.Cells(counter,108).Value = item # 108 is column DD
                    counter += 1
                counter = 2
                for item in CableTypeList:
                    Export.ActiveSheet.Cells(counter,109).Value = item # 109 is column DE
                    counter += 1
                counter = 2
                for item in ScreenList:
                    Export.ActiveSheet.Cells(counter,110).Value = item # 110 is column DF
                    counter += 1
                 #The following commands are apparently necessary to insert an updated drop-down list in Excel
                xlUp = -4162
                LastRow = Export.ActiveSheet.Cells(Export.ActiveSheet.Rows.Count, "A").End(xlUp).Row
                print("Last row in Cable Export: ",LastRow)

                """
                As the object is a cable, the cable parameters can be filled in from the
                CableTypeDict based on the model number. The format will be:
                CableTypeDict[obj.Model] = ['Belden', '973107Z', '2Pr. 22 AWG OS', '300V', 
                        '22 AWG', 'LSZH-TS', '15.2mm', 'M20', None, None, None, None, '', 2, 2, 'OS']
                Assignment:
                    'Belden'        : CableTypeDict[obj.Model][0] : Manufacturer        : column H=8
                    '973107Z'       : CableTypeDict[obj.Model][1] : Not Assigned (yet)
                    '2Pr. 22 AWG OS': CableTypeDict[obj.Model][2] : CoreConfiguration   : column AT=46
                    '300V'          : CableTypeDict[obj.Model][3] : VoltageRating       : column AX=50
                    '22 AWG'        : CableTypeDict[obj.Model][4] : CrossSection        : column AN=40
                    'LSZH-TS'       : CableTypeDict[obj.Model][5] : FlameRetardancy     : column AK=37
                    '15.2mm'        : CableTypeDict[obj.Model][6] : OverallDiameter     : column AQ=43
                    'M20'           : CableTypeDict[obj.Model][7] : CableAssembly       : column AM=39
                    None            : CableTypeDict[obj.Model][8] : ChargingCurrent     : column BE=57
                    None            : CableTypeDict[obj.Model][9] : Resistance          : column AY=51
                    None            : CableTypeDict[obj.Model][10]: Capacitance         : column AZ=52
                    None            : CableTypeDict[obj.Model][11]: Weight              : column AR=44
                """
                for row in range(2,LastRow+1):
                    Model = Export.ActiveSheet.Cells(row,9).Value                   
                    if Model in CableTypeDict:
                        Export.ActiveSheet.Cells(row,8).Value = CableTypeDict[Model][0]   # Manufacturer
                        Export.ActiveSheet.Cells(row,46).Value = CableTypeDict[Model][2]  # CoreConfiguration
                        Export.ActiveSheet.Cells(row,49).Value = CableTypeDict[Model][3]  # VoltageRating
                        Export.ActiveSheet.Cells(row,40).Value = CableTypeDict[Model][4]  # CrossSection
                        Export.ActiveSheet.Cells(row,37).Value = CableTypeDict[Model][5]  # FlameRetardancy
                        Export.ActiveSheet.Cells(row,43).Value = CableTypeDict[Model][6]  # OverallDiameter
                        Export.ActiveSheet.Cells(row,39).Value = CableTypeDict[Model][7]  # CableAssembly
                        Export.ActiveSheet.Cells(row,57).Value = CableTypeDict[Model][8]  # ChargingCurrent
                        Export.ActiveSheet.Cells(row,50).Value = CableTypeDict[Model][9]  # Resistance
                        Export.ActiveSheet.Cells(row,51).Value = CableTypeDict[Model][10] # Capacitance
                        Export.ActiveSheet.Cells(row,44).Value = CableTypeDict[Model][11] # Weight
                for row in range(2,LastRow+1):
                   # Construction of the updated drop-down list
                    Export.ActiveSheet.Cells(row,9).Validation.Add(3, 1, 3, Range2)   # column I=9, caption:Model
                    Export.ActiveSheet.Cells(row,9).Validation.InCellDropdown = True  
                    Export.ActiveSheet.Cells(row,47).Validation.Add(3, 1, 3, Range3)  # column AU=47, caption: ScreenConfiguration
                    Export.ActiveSheet.Cells(row,47).Validation.InCellDropdown = True                     
                    Export.ActiveSheet.Cells(row,48).Validation.Add(3, 1, 3, Range)   # column AV=48, caption: Connection1
                    Export.ActiveSheet.Cells(row,48).Validation.InCellDropdown = True                  
                    Export.ActiveSheet.Cells(row,49).Validation.Add(3, 1, 3, Range)   # column AW=49, caption: Connection2
                    Export.ActiveSheet.Cells(row,49).Validation.InCellDropdown = True              
#---------------------------------------------------------------------------------------------            
        Export.ActiveSheet.Range("A:EE").Columns.AutoFit()
        Export.Close(SaveChanges=True)
        Xcel.Application.Quit()
        progress.stop()
        connection.close()
        storage.close()
        text = "Object Export prepared.\nDo you want to save the Export file\nin the Object Exports directory?"
        answer = tk.messagebox.askokcancel(title=None, message=text)
        print(answer)
        if answer:
            src = path
            dst = os.getcwd()+"\Object Exports"+"\Object Export.xlsx"
            shutil.copyfile(src,dst)
            now = datetime.now()
            dt_string = now.strftime("%d-%m-%Y %H-%M-%S")
            filename = chosen_object+"-"+dt_string+".xlsx"
            print(filename)
            old_name = dst
            new_name = os.getcwd()+"\Object Exports"+"\\"+filename
            os.rename(old_name, new_name)
            tk.messagebox.showwarning(title=None, message="Object Export saved\nin Object Exports")
        pythoncom.CoUninitialize()
        btn_export.config(state=tk.NORMAL)
        btn_import.config(state=tk.NORMAL)
    except BaseException as e:
       print(e.args)
       Xcel.Application.Quit()
       connection.close()
       storage.close()
       progress.stop()
       pythoncom.CoUninitialize()
       btn_export.config(state=tk.NORMAL)
       btn_import.config(state=tk.NORMAL)        

def startimport():
    print("STARTIMPORT")
    threading.Thread(target=importation).start()        
      
def importation():
    print("IMPORT")
    global addr
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()
    pythoncom.CoInitialize()
    try:
        btn_export.config(state=tk.DISABLED)
        btn_import.config(state=tk.DISABLED)        
        ATTR_List = []
        TAG_List = []
        path = os.getcwd()+"\Object Export.xlsx"
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Import = Xcel.Workbooks.Open(path)       
        xlToLeft = -4159
        xlUp = -4162
        LastCol = Import.ActiveSheet.Cells(1, Import.ActiveSheet.Columns.Count).End(xlToLeft).Column
        LastRow = Import.ActiveSheet.Cells(Import.ActiveSheet.Rows.Count, "A").End(xlUp).Row
        for column in range(1,LastCol+1): 
            ATTR_List.append(Import.ActiveSheet.Cells(1,column).Value)
            progress.update()
        for row in range(2,LastRow+1):
            TAG_List.append(Import.ActiveSheet.Cells(row,1).Value)
            progress.update()
#        print(LastRow,LastCol)
#        print(ATTR_List)
#        print(TAG_List)
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename) 
        storage = ClientStorage.ClientStorage(addr)           
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        row = 2
        for tag in TAG_List:
            LocalObject = root[tag]
            column = 1
            for attribute in ATTR_List:
                CellValue = Import.ActiveSheet.Cells(row,column).Value
                setattr(LocalObject, attribute, CellValue)
                column += 1
                progress.update()
            row += 1
            """
            In case of a cable, the cable colors can be added
            based on the CoreTypeDict compiled during startup.
            The format is a list of tuples like this:
            'BFOU(c)4P-0.75': [('1', '1', '1BK'), ('1', '2', '1BU'), 
            ('2', '3', '2BK'), ('2', '4', '2BU'), ('3', '5', '3BK'), 
            ('3', '6', '3BU'), ('4', '7', '4BK'), ('4', '8', '4BU'), 
            ('5', '9', 'OAS')]
            """
            if isinstance(LocalObject, DS_BES_Cable):
                if LocalObject.Model in CoreTypeDict:
                    for i in range(0,len(CoreTypeDict[LocalObject.Model])):                           
                        setattr(LocalObject,"Color"+str(i+1),CoreTypeDict[LocalObject.Model][i][-1])
                        setattr(LocalObject,"CoreNumber"+str(i+1),str(CoreTypeDict[LocalObject.Model][i][0]))
            root[tag] = LocalObject
            transaction.commit()         
        connection.close()
        storage.close()
        Import.Close(SaveChanges=True)
        Xcel.Application.Quit()
        progress.stop()
        pythoncom.CoUninitialize()
        tk.messagebox.showwarning(title=None, message="Objects Import finished.")
        btn_export.config(state=tk.NORMAL)
        btn_import.config(state=tk.NORMAL)       
    except BaseException as e:
        print(e.args)
        connection.close()
        storage.close()
        pythoncom.CoUninitialize()
        btn_export.config(state=tk.NORMAL)
        btn_import.config(state=tk.NORMAL)

def insertAcad():
    print("INSERTACAD")
    global addr
    selection = []
    items = lbox_objects.curselection()
    for item in items:
        op = lbox_objects.get(item)
        selection.append(op)
    if len(selection) == 0:
        tk.messagebox.showwarning(title=None, message="No objects selected.")
        return
    try:
        def vtPnt(x, y, z=0.0):
            # Convert coordinate points to floating point numbers
            return win32com.client.VARIANT(pythoncom.VT_ARRAY | pythoncom.VT_R8, (x, y, z))
        directory = os.getcwd()+"/ACAD"
        filename = "/" + "TEMPLATE.dwg" 
        acad = win32com.client.Dispatch("Autocad.Application")
        doc = acad.Documents.Open(directory+filename)
        time.sleep(2)
#        color = acad.GetInterfaceObject("AutoCAD.AcCmColor.24")       
#        color.SetRGB(255,255,0)
        acad.Visible = False
        progress['mode'] = 'indeterminate'
        progress.start()
        progress.update()
        time.sleep(1)
        # First collect Tags
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr) 
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        # determine class type of selected objects
        # placed in a dictionary to keep the database
        # open as short as possible
        BlockStringDict = {}
        for item in selection:
            if isinstance(root[item], DS_BES_Antenna):
                BlockStringDict[item]  = "ANTENNA"            
            if isinstance(root[item], DS_BES_Beacon):
                BlockStringDict[item]  = "ALARM BEACON"
            elif isinstance(root[item], DS_BES_Cable):
                BlockStringDict[item]  = "TEXT"
            elif isinstance(root[item], DS_BES_CV):
                BlockStringDict[item]  = "TEXT"
            elif isinstance(root[item], DS_BES_Enclosure):
                BlockStringDict[item]  = "TEXT-POLYLINE"  
            elif isinstance(root[item], DS_BES_Handsw):
                BlockStringDict[item]  = "SWITCH"  
            elif isinstance(root[item], DS_BES_Limitsw):
                BlockStringDict[item]  = "SWITCH"  
            elif isinstance(root[item], DS_BES_SOL_V):
                BlockStringDict[item]  = "TEXT"  
            elif isinstance(root[item], DS_BES_Transformer):
                BlockStringDict[item]  = "TEXT"
            else:
                BlockStringDict[item] = "LOCAL MOUNTED INSTRUMENT"
            progress.update()
        print(BlockStringDict)
        connection.close()
        storage.close()
        horizontal = 50.0
        vertical = -30.0
        counter = 1           
        for item in selection:
            centerPoint = vtPnt(horizontal, vertical)
            if "TEXT" in BlockStringDict[item]: 
                Text = acad.ActiveDocument.ModelSpace.AddText(item,centerPoint,3)
#                Text.TrueColor = color
            else:
                Block = acad.ActiveDocument.ModelSpace.InsertBlock(centerPoint,BlockStringDict[item],1,1,1,0)
                for attrib in Block.GetAttributes():
                    if attrib.TagString == "SC":
                       attrib.TextString = item.split("-")[0]
                    if attrib.TagString == "IF":
                       attrib.TextString = item.split("-")[1]
                    if attrib.TagString == "TN":
                       attrib.TextString = item.split("-")[2]
            horizontal += 50
            counter += 1
            if counter == 6:
                horizontal = 50
                vertical -= 30.0
                counter = 1
            progress.update()
        doc.Close()
        acad.Application.Quit()
        del acad
        progress.stop()
        if "DADispatcherService.exe" in (p.name() for p in psutil.process_iter()):
            os.system("taskkill /f /im  DADispatcherService.exe")
        tk.messagebox.showwarning(title=None, message="AutoCad Exportation finished.")     
    except BaseException as e:
        print(e.args)
        doc.close()
        acad.Application.Quit()
        progress.stop()
        if "DADispatcherService.exe" in (p.name() for p in psutil.process_iter()):
            os.system("taskkill /f /im  DADispatcherService.exe")
            
def objectnumbers():
    print("OBJECTNUMBERS")
    global OBJNumbersDict
    txt1, txt2 = "",""
    for key in OBJNumbersDict:
        txt1 = txt1 + key + "\n"
        txt2 = txt2 + str(OBJNumbersDict[key]) + "\n"

    x = window.winfo_x()
    y = window.winfo_y()
    top = tk.Toplevel(window)
    lbl_objectnames = tk.Label(master=top,text =txt1,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_objectnames.grid(row=0, column=0, sticky="nsew")
    lbl_objectnumbers = tk.Label(master=top,text =txt2,justify=tk.LEFT,relief=tk.GROOVE)
    lbl_objectnumbers.grid(row=0, column=1, sticky="nsew")    
    top.geometry("+%d+%d" % (x + 900, y + 200))       
    top.mainloop()            
        
def startPCexport():
    threading.Thread(target=PCexportation).start() 
     
def PCexportation():
    print("PCEXPORTATION")
    pythoncom.CoInitialize()
    global OBJ_List, addr
    try:
#        print(OBJ_List)
#        print(ObjectChoice.get())
        btn_PCexport.config(state=tk.DISABLED)
        btn_PCimport.config(state=tk.DISABLED)
        progress['mode'] = 'indeterminate'
        progress.start()
        progress.update()
        if lbl_objectchoice['text'] == "ALL":
            tk.messagebox.showwarning(title=None, message="First Select a Specific Object Type.")
            progress.stop()
            btn_PCexport.config(state=tk.NORMAL)
            btn_PCimport.config(state=tk.NORMAL)
            return
        LocalObject = None
        if (ObjectChoice.get() == "ANALOG INPUTS" or ObjectChoice.get() == "ANALOG OUTPUTS" or
            ObjectChoice.get() == "ANTENNAS" or ObjectChoice.get() == "BEACONS" or ObjectChoice.get() == "CABLES" or 
            ObjectChoice.get() == "COMPASSES" or ObjectChoice.get() == "DIGITAL INPUTS" or ObjectChoice.get() == "DIGITAL OUTPUTS" or 
            ObjectChoice.get() == "ENCLOSURES" or ObjectChoice.get() == "FIRE DETECTORS" or ObjectChoice.get() == "FOGDETECTORS" or 
            ObjectChoice.get() == "FOGHORNS" or ObjectChoice.get() == "GAS DETECTORS" or ObjectChoice.get() == "HANDSWITCHES" or 
            ObjectChoice.get() == "LIMIT SWITCHES" or ObjectChoice.get() == "LOADPINS" or ObjectChoice.get() == "OCEANOGRAPHS" or 
            ObjectChoice.get() == "PIG DETECTORS" or ObjectChoice.get() == "SLIPRINGS" or ObjectChoice.get() == "SPEAKERS" or 
            ObjectChoice.get() == "SOLARPANELS" or ObjectChoice.get() == "WEATHERSTATIONS" or ObjectChoice.get() == "WINDGENERATORS" or 
            ObjectChoice.get() == "TRANSFORMERS"):
            tk.messagebox.showwarning(title=None, message="Selected object does not have process conditions.")
            progress.stop()
            btn_PCexport.config(state=tk.NORMAL)
            btn_PCimport.config(state=tk.NORMAL)
            return
            LocalObject = None            
        if ObjectChoice.get() == "CONTROL VALVES":
            LocalObject = DS_BES_CV()
        if ObjectChoice.get() == "FLOWMETERS":
            LocalObject = DS_BES_FI()
        if ObjectChoice.get() == "LEVEL GAUGES":
            LocalObject = DS_BES_LG()
        if ObjectChoice.get() == "LEVEL TRANSMITTERS":
            LocalObject = DS_BES_LIT()
        if ObjectChoice.get() == "PRESSURE GAUGES":
            LocalObject = DS_BES_PG()
        if ObjectChoice.get() == "PRESSURE TRANSMITERS":
            LocalObject = DS_BES_PIT()
        if ObjectChoice.get() == "RESTRICTION ORIFICES":
            LocalObject = DS_BES_RO()              
        if ObjectChoice.get() == "SOLENOIDS":
            LocalObject = DS_BES_SOL_V()           
        if ObjectChoice.get() == "SAFETY VALVES":
            LocalObject = DS_BES_SV()
        if ObjectChoice.get() == "TEMPERATURE GAUGES":
            LocalObject = DS_BES_TG()
        if ObjectChoice.get() == "TEMPERATURE TRANSMITTERS":
            LocalObject = DS_BES_TT()                       
        path = os.getcwd()+"\ObjectPC Export.xlsx"
        # Remove exisiting collection and create new one
        if os.path.exists(path):
            os.remove(path)
        time.sleep(2) # To give the os module time to remove the Excel Export File
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Export = Xcel.Workbooks.Add() 
        Export.SaveAs(path)
        Xcel.Visible = False
        # Add the headers
        column = 1
        for key in LocalObject.FieldsDict:  #######
            if (key == 'TagNumber' or key =='ServiceDescription' or key == 'PandID' or key == 'VesselNumber' or key == 'LineID' or key == 'Size' or key == 'Schedule'):
                Export.ActiveSheet.Cells(1,column).Value = key
                Export.ActiveSheet.Cells(1,column).Font.Bold = True
                column = column + 1
                progress.update()
        for key in LocalObject.InternalFieldsDict:
             if  ('_PC') in LocalObject.InternalFieldsDict[key]  :
                 Export.ActiveSheet.Cells(1,column).Value = key
             Export.ActiveSheet.Cells(1,column).Font.Bold = True
             column = column + 1
             progress.update()
        row = 2
        column = 1
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)        
        db = DB(storage)
        connection = db.open()
        root = connection.root()     
        for item in OBJ_List:
            LocalObject = root[item]
            for key in LocalObject.FieldsDict:
                if (key == 'TagNumber' or key == 'ServiceDescription' or key == 'PandID' or key == 'VesselNumber' or key == 'LineID' or key == 'Size' or key == 'Schedule'):
                     Export.ActiveSheet.Cells(row,column).Value = getattr(LocalObject, key)
                     column = column + 1
            for key in LocalObject.InternalFieldsDict:
                 if '_PC' in LocalObject.InternalFieldsDict[key]:
                     Export.ActiveSheet.Cells(row,column).Value = getattr(LocalObject, key)
                 column = column + 1
            row = row + 1
            column = 1
            progress.update()
        connection.close()
        storage.close()
#---------------------------------------------------------------------------------------------            
        Export.ActiveSheet.Range("A:EE").Columns.AutoFit()
        Export.Close(SaveChanges=True)
        Xcel.Application.Quit()
        progress.stop()        
        tk.messagebox.showwarning(title=None, message="ObjectPC Export prepared.")
        connection.close()
        storage.close()
        pythoncom.CoUninitialize()
        btn_PCexport.config(state=tk.NORMAL)
        btn_PCimport.config(state=tk.NORMAL)
    except BaseException as e:
       print(e.args)
       Xcel.Application.Quit()
       connection.close()
       storage.close()
       progress.stop()
       pythoncom.CoUninitialize()
       btn_PCexport.config(state=tk.NORMAL)
       btn_PCimport.config(state=tk.NORMAL)

def startPCimport():
    print("STARTPCIMPORT")
    threading.Thread(target=PCimportation).start()        
      
def PCimportation():
    print("PCIMPORT")
    progress['mode'] = 'indeterminate'
    progress.start()
    progress.update()
    pythoncom.CoInitialize()
    try:
        btn_PCexport.config(state=tk.DISABLED)
        btn_PCimport.config(state=tk.DISABLED)        
        ATTB_List = []
        TAG_List = []
        ATTB_Dict={}
        path = os.getcwd()+"\ObjectPC Export.xlsx"
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Import = Xcel.Workbooks.Open(path)       
        xlToLeft = -4159
        xlUp = -4162
        LastCol = Import.ActiveSheet.Cells(1, Import.ActiveSheet.Columns.Count).End(xlToLeft).Column
        LastRow = Import.ActiveSheet.Cells(Import.ActiveSheet.Rows.Count, "A").End(xlUp).Row
        for column in range(2,LastCol+1):
            ATTB_List.append(Import.ActiveSheet.Cells(1,column).Value)
            progress.update()
        for row in range(2,LastRow+1):
            TAG_List.append(Import.ActiveSheet.Cells(row,1).Value)
            progress.update()
        row = 2
        for Tag in TAG_List:
            Value_list=[]
            for column in range(2,LastCol+1):
                Value_list.append(Import.ActiveSheet.Cells(row,column).Value)
                ATTB_Dict[Tag]=Value_list
                progress.update()
            row+=1  
        Import.Close(SaveChanges=True)
        Xcel.Application.Quit()
#        print(LastRow,LastCol)
        print(ATTB_Dict)
#        print(TAG_List)
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)              
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for key in ATTB_Dict:
            LocalObject = root[key]
            for attribute in ATTB_List:
                setattr(LocalObject, attribute, ATTB_Dict[key][ATTB_List.index(attribute)])
                progress.update()
            root[key] = LocalObject
            transaction.commit()            
        connection.close()
        storage.close()         
        progress.stop()
        pythoncom.CoUninitialize()
        tk.messagebox.showwarning(title=None, message="ObjectsPC import finished")
        btn_PCexport.config(state=tk.NORMAL)
        btn_PCimport.config(state=tk.NORMAL)       
    except BaseException as e:
        print(e.args)
        connection.close()
        storage.close()
        pythoncom.CoUninitialize()
        btn_PCexport.config(state=tk.NORMAL)
        btn_PCimport.config(state=tk.NORMAL)       
   
def collect():
    print("COLLECT")
    global addr
    try:
        btn_collect.config(state=tk.DISABLED)
        if lbl_objectchoice['text'] == "ALL":
            tk.messagebox.showwarning(title=None, message="First Select a Specific Object Type.")
            progress.stop()
            btn_collect.config(state=tk.NORMAL)
            return
        LocalObject  = None
        if ObjectChoice.get() == "ANALOG INPUTS":
            LocalObject = DS_BES_AI
        if ObjectChoice.get() == "ANALOG OUTPUTS":
            LocalObject = DS_BES_AO        
        if ObjectChoice.get() == "ANTENNAS":
            LocalObject = DS_BES_Antenna
        if ObjectChoice.get() == "BEACONS":
            LocalObject = DS_BES_Beacon
        if ObjectChoice.get() == "CABLES":
            LocalObject = DS_BES_Cable
        if ObjectChoice.get() == "COMPASSES":  
            LocalObject = DS_BES_Compass
        if ObjectChoice.get() == "CONTROL VALVES":
            LocalObject = DS_BES_CV
        if ObjectChoice.get() == "DIGITAL INPUTS":
            LocalObject = DS_BES_DI
        if ObjectChoice.get() == "DIGITAL OUTPUTS":
            LocalObject = DS_BES_DO            
        if ObjectChoice.get() == "ENCLOSURES":
            LocalObject = DS_BES_Enclosure
        if ObjectChoice.get() == "FIRE DETECTORS":
            LocalObject = DS_BES_FD
        if ObjectChoice.get() == "FLOWMETERS":
            LocalObject = DS_BES_FI
        if ObjectChoice.get() == "FOGDETECTORS":
            LocalObject = DS_BES_Fogdetector
        if ObjectChoice.get() == "FOGHORNS":    
            LocalObject = DS_BES_Foghorn
        if ObjectChoice.get() == "GAS DETECTORS":
            LocalObject = DS_BES_GD
        if ObjectChoice.get() == "HANDSWITCHES":
            LocalObject = DS_BES_Handsw
        if ObjectChoice.get() == "LEVEL GAUGES":
            LocalObject = DS_BES_LG
        if ObjectChoice.get() == "LIMIT SWITCHES":
            LocalObject = DS_BES_Limitsw
        if ObjectChoice.get() == "LEVEL TRANSMITTERS":
            LocalObject = DS_BES_LIT
        if ObjectChoice.get() == "LOADPINS": 
            LocalObject = DS_BES_Loadpin
        if ObjectChoice.get() == "OCEANOGRAPHS":
            LocalObject = DS_BES_Oceanograph 
        if ObjectChoice.get() == "PRESSURE GAUGES":
            LocalObject = DS_BES_PG
        if ObjectChoice.get() == "PIG DETECTORS":
            LocalObject = DS_BES_PIG
        if ObjectChoice.get() == "PRESSURE TRANSMITERS":
            LocalObject = DS_BES_PIT
        if ObjectChoice.get() == "RESTRICTION ORIFICES":
            LocalObject = DS_BES_RO
        if ObjectChoice.get() == "SLIPRINGS":
            LocalObject = DS_BES_Slipring
        if ObjectChoice.get() == "SPEAKERS":
            LocalObject = DS_BES_Speaker            
        if ObjectChoice.get() == "SOLENOIDS":
            LocalObject = DS_BES_SOL_V
        if ObjectChoice.get() == "SOLARPANELS":
            LocalObject = DS_BES_Solarpanel
        if ObjectChoice.get() == "SAFETY VALVES":
            LocalObject = DS_BES_SV
        if ObjectChoice.get() == "TEMPERATURE GAUGES":
            LocalObject = DS_BES_TG
        if ObjectChoice.get() == "TEMPERATURE TRANSMITTERS":
            LocalObject = DS_BES_TT
        if ObjectChoice.get() == "WEATHERSTATIONS":
            LocalObject = DS_BES_Weatherstation
        if ObjectChoice.get() == "WINDGENERATORS":
            LocalObject = DS_BES_Windgenerator
        if ObjectChoice.get() == "TRANSFORMERS":
            LocalObject = DS_BES_Transformer          
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr)
        db = DB(storage)
        connection = db.open()
        path1 = os.getcwd()+"\Datasheets\\"
        path2 = os.getcwd()+"\Collection.xlsx"
        # Remove exisiting collection and create new one
        if os.path.exists(path2):
            os.remove(path2)
        time.sleep(2)
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Collection = Xcel.Workbooks.Add()
        Collection.SaveAs(path2)
        counter = 0
        for item in OBJ_List:
            Object = root[OBJ_List[counter]]
            ObjectName = LocalObject.__name__ + ".xlsx"
            print(ObjectName, LocalObject.__name__)
            path = path1 + ObjectName
            Datasheet = Xcel.Workbooks.Open(path)
            for key in LocalObject.InternalFieldsDict:
                Datasheet.ActiveSheet.Range(Object.InternalFieldsDict[key]).Value = getattr(Object, key)
            for key in LocalObject.FieldsDict:
                Datasheet.ActiveSheet.Range(Object.FieldsDict[key]).Value = getattr(Object, key)
            DataWorksheet = Datasheet.Worksheets(1)
            DataWorksheet.Name = Object.TagNumber          
            DataWorksheet.Copy(After=Collection.Sheets(Collection.Sheets.Count))
            Datasheet.Close(SaveChanges=False)
            counter = counter + 1
        Collection.Close(SaveChanges=True)        
        connection.close()
        storage.close()
        Xcel.Application.Quit()
        os.system('TASKKILL /F /IM excel.exe')
        tk.messagebox.showwarning(title=None, message="Collection prepared.")
        btn_collect.config(state=tk.NORMAL)
    except BaseException as e: 
        print(e.args)
        connection.close()
        storage.close()
        os.system('TASKKILL /F /IM excel.exe')
        btn_collect.config(state=tk.NORMAL)    
    
def cableschedule():
    print("CABLESCHEDULE")
    global addr
    try:
        btn_Cableschedule.config(state=tk.DISABLED)
        progress.start()
        CableList = []
        CableDict = {}
        ObjectDict = {}
        ScheduleLine = []
        ScheduleList = []
        LocalObject = None
#        directory = os.getcwd()+"/Databases"
#        filename = "/Lufeng.fs"
#        storage = FileStorage.FileStorage(directory+filename)
        storage = ClientStorage.ClientStorage(addr) 
        db = DB(storage)
        connection = db.open()
        root = connection.root()
        for item in root:
            LocalObject = root[item]
            if isinstance(LocalObject, DS_BES_Cable):
                CableDict['TagNumber'] = LocalObject.TagNumber
                CableDict['FlameRetardancy'] = LocalObject.FlameRetardancy                
                CableDict['VoltageRating'] = LocalObject.VoltageRating
                CableDict['CoreConfiguration'] = LocalObject.CoreConfiguration
                CableDict['OverallDiameter'] = LocalObject.OverallDiameter
                CableDict['Connection1'] = LocalObject.Connection1
                CableDict['Connection2'] = LocalObject.Connection2
                CableList.append(CableDict)
                CableDict = {}
#            else:
#                if not isinstance(LocalObject, DS_BES_Slipring):
#                    ObjectDict[LocalObject.TagNumber] = []
#                    ObjectDict[LocalObject.TagNumber].append(LocalObject.ServiceDescription)
#                    ObjectDict[LocalObject.TagNumber].append(LocalObject.CabinetLocation)
#                if isinstance(LocalObject, DS_BES_Slipring):
#                    ObjectDict[LocalObject.TagNumber] = []
#                    ObjectDict[LocalObject.TagNumber].append(LocalObject.ServiceDescription)
#                    ObjectDict[LocalObject.TagNumber].append("Turret")                    

        # print("HERE",len(CableList))
        connection.close()
        storage.close()
        for dictionary in CableList:
            # print(dictionary)
#            print(ObjectDict)
            ScheduleLine = []
            ScheduleLine.append(dictionary['TagNumber'])
            ScheduleLine.append('') # In a later stage to be filled with Length of the cable
            if dictionary['FlameRetardancy'] != None and dictionary['VoltageRating'] != None:
                Column4 = dictionary['FlameRetardancy']+" ("+dictionary['VoltageRating']+")"
            elif dictionary['FlameRetardancy'] != None:
                Column4 = dictionary['FlameRetardancy']                
            else:
                Column4 = ""
            ScheduleLine.append(Column4)
            ScheduleLine.append(dictionary['CoreConfiguration'])
            ScheduleLine.append(dictionary['OverallDiameter'])
            GlandSize = ''
            if dictionary['OverallDiameter'] == None or dictionary['OverallDiameter'] == '':
                GlandSize = ''
            elif not (type(dictionary['OverallDiameter']) is str):
                if 6.1 <= float(dictionary['OverallDiameter']) <= 13.1:
                    GlandSize = 'M16'                
                if 13.1 <= float(dictionary['OverallDiameter']) <= 20.9:
                    GlandSize = 'M20'                
                if 20.9 <= float(dictionary['OverallDiameter']) <= 26.2:
                    GlandSize = 'M25'
                if 26.2 <= float(dictionary['OverallDiameter']) <= 33.9:
                    GlandSize = 'M32'
                if 33.9 <= float(dictionary['OverallDiameter']) <= 40.4:
                    GlandSize = 'M40'
                if 40.4 <= float(dictionary['OverallDiameter']) <= 53.0:
                    GlandSize = 'M50'
            ScheduleLine.append(dictionary['Connection1'])
            ScheduleLine.append('') # in a later stage to be filled with a description of Connection1
            ScheduleLine.append('') # In a later stage to be filled with the location of Connection1
            ScheduleLine.append(GlandSize)
            ScheduleLine.append(dictionary['Connection2'])
            ScheduleLine.append('') # in a later stage to be filled with a description of Connection2
            ScheduleLine.append('') # In a later stage to be filled with the location of Connection2            
            ScheduleLine.append(GlandSize)            
            ScheduleList.append(ScheduleLine)

        # print("HERE",len(ScheduleList)  )
        
        print('opening excel')
        Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
        Xcel.Visible = False   
        path = os.getcwd()+"\Cableschedule.xlsx"
#        if os.path.exists(path):
#            os.remove(path)
        Cableschedule = Xcel.Workbooks.Open(path)
        # first clean the Cableschedule template
        xlUp = -4162
        LastRow = Cableschedule.ActiveSheet.Cells(Cableschedule.ActiveSheet.Rows.Count, "B").End(xlUp).Row
        column = 1
        for row in range(11,LastRow+1):
            for column in range(1,16):
                Xcel.ActiveSheet.Cells(row,column).Value = ""
            row += 1
        ScheduleList.sort()
        row = 11
        column = 2
        for Line in ScheduleList:
            print(Line)
            for item in Line:
                Xcel.ActiveSheet.Cells(row,column).Value = item
                column += 1
                progress.update()
            column = 2
            row += 1
        # Xcel.ActiveSheet.Range("A:P").Rows.RowHeight = 26
        Cableschedule.ActiveSheet.Range("A:O").Columns.AutoFit()
        Cableschedule.Close(SaveChanges=True)
        Xcel.Application.Quit()       
        btn_Cableschedule.config(state=tk.NORMAL)
        tk.messagebox.showwarning(title=None, message="Cableschedule finished.")
        progress.stop()       
    except BaseException as e: 
        print(e.args)
        connection.close()
        storage.close()
        os.system('TASKKILL /F /IM excel.exe')
        btn_Cableschedule.config(state=tk.NORMAL)
        progress.stop()

def iolist():
    print("IOLIST")
    print('opening excel')
    Xcel = win32com.client.gencache.EnsureDispatch("Excel.Application")
    Xcel.Visible = False   
    path = os.getcwd()+"\IOList.xlsx"
    IOList = Xcel.Workbooks.Open(path)
    IOList.Worksheets("ANALOG SIGNALS").Cells(3,5).Value = "110"
    IOList.Worksheets("DIGITAL SIGNALS").Cells(3,5).Value = "110"
    IOList.Worksheets("SERIAL INTEFACE SIGNALS").Cells(2,1).Value = "WHATEVER"
    IOList.Close(SaveChanges=True)
    Xcel.Application.Quit()
    
def enclosureterminations():
    print("ENCLOSURE TERMINATIONS")
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
        btn_Terminations.config(state=tk.DISABLED)
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
            CableDict[key][153] = Start of Color1..50
            CableDict[key][203] Start of RSCR1..50
            CableDict[key][253] = Start of RTermination1..50
            CableDict[key][303] Start of RSignal1..50
            CableDict[key][353] = Start of CoreNumber1..50
            CableDict[key][403] CoreConfiguration e.g. '2Pr. 22 AWG OS'   
            obtain the cables, the connections and the number of cores per cable in a list per cable
            list comprehension:        
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
                for i in range(1,51): # maximum number of cores is 41
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
                for i in range(1,51): # maximum number of cores is 41
                    CableDict[key].append(getattr(obj,"RSignal"+str(i)))
                for i in range(1,51): # maximum number of cores is 41
                    CableDict[key].append(getattr(obj,"CoreNumber"+str(i)))                    
                if obj.CoreConfiguration != None:
                    CableDict[key].append(obj.CoreConfiguration)
                else:
                    CableDict[key].append("Unknown")
                """
                All these for statements to make sure we get a proper ordening of attributes in the list:
                [Connection1, Coreconfiguration, Core Number, Connection2,
                 LSignal1..51, LTermination1..51, LSCR1..51, LCore1..51,
                 RCore1..51, RSCR1..51, RTermination1..51, RSignal1..51] 
                """
            # Next compile EnclosureDict
            # This contains the Enclosures with for each enclosure all cables attached to it.
            # Format EnclosureTag: [ Cable 1, Cable2,....,SPARE,SPARE,..]            
            if isinstance(obj,DS_BES_Enclosure):
                #print(obj.__dict__)                
                EnclosureDict[key] = []
                for i in range(1,31):
                    attribute = 'Connection'+str(i)
                    if attribute in obj.__dict__:
                        #print(attribute,obj.__dict__[attribute])
                        if obj.__dict__[attribute] != "" or obj.__dict__[attribute] != None:
                            EnclosureDict[key].append(obj.__dict__[attribute])
                # Clean EquipmentDict[key] from empty strings
                EnclosureDict[key] = [i for i in EnclosureDict[key] if i]                                    
        connection.close()
        storage.close()
        progress.stop()
        print(len(CableDict),len(EnclosureDict))
        #Sort EnclosureDict
        EnclosureTuples = sorted(EnclosureDict.items())
        EnclosureDict = dict(EnclosureTuples)
        # Export the Enclosures separately to an Excel Workbook
        path = os.getcwd()+"\EnlosureTerminations.xlsm"
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
        HandledCables = []
##        print(CableDict)
#        print(EnCab)
        for key in EnclosureDict:
            progress.update()
            row = 1
            Terminations.Worksheets.Add(After = Terminations.Worksheets(Terminations.Worksheets.Count))
            Terminations.ActiveSheet.Name = key
            startrow = row
            Terminations.ActiveSheet.Cells(row,1).Value = key
            Terminations.ActiveSheet.Range("A"+str(row)+":"+"H"+str(row)).Merge() 
            Terminations.ActiveSheet.Range("A"+str(row)+":"+"H"+str(row)).Interior.ColorIndex = 15 
            Terminations.ActiveSheet.Range("A"+str(row)+":"+"H"+str(row)).Font.Bold = True
            row += 1
            for item in EnclosureDict[key]:
                progress.update()
#                if item == "SPARE":
#                    startrow2 = row
#                    Terminations.ActiveSheet.Cells(row,9).Value = Header2[0]
#                    Terminations.ActiveSheet.Cells(row,10).Value = Header2[1]
#                    Terminations.ActiveSheet.Cells(row,11).Value = Header2[2]
#                    Terminations.ActiveSheet.Cells(row,12).Value = item # cable tag
#                    Terminations.ActiveSheet.Cells(row,12).Interior.ColorIndex = 40
#                    Terminations.ActiveSheet.Cells(row,13).Value = Header2[4]
#                    Terminations.ActiveSheet.Cells(row,14).Value = Header2[5]
#                    Terminations.ActiveSheet.Cells(row,15).Value = Header2[6]
#                    row+=1
                if item != "SPARE":
                    startrow2 = row
                    cable = item
                    Terminations.ActiveSheet.Cells(row,1).Value = Header2[0]
                    Terminations.ActiveSheet.Cells(row,2).Value = Header2[1]
                    Terminations.ActiveSheet.Cells(row,3).Value = Header2[2]
                    Terminations.ActiveSheet.Cells(row,4).Value = cable # cable tag
                    Terminations.ActiveSheet.Cells(row,4).Interior.ColorIndex = 4
                    Terminations.ActiveSheet.Cells(row,5).Value = Header2[4]
                    Terminations.ActiveSheet.Cells(row,6).Value = Header2[5]
                    Terminations.ActiveSheet.Cells(row,7).Value = Header2[6]
                    condition1 = CableDict[item][2] != key
                    condition2 = CableDict[item][2] == key
                    if condition1:
                        Terminations.ActiveSheet.Cells(row,8).Value = CableDict[cable][2]
                    if condition2:
                        Terminations.ActiveSheet.Cells(row,8).Value = CableDict[cable][0]
                    Terminations.ActiveSheet.Cells(row,8).Interior.Color = rgbToInt((226,239,218))
                    row += 1
                    if item not in HandledCables:
                        prefix = "bb"
                    else:
                        prefix = "cc"                
                    if type(CableDict[item][1]) is list:
                        Name = prefix+item.replace('/','') # remove forwardslashes from cable names
                        Name = Name.replace('-','') # remove dashes from cable names
                        if CableDict[item][1][2] == 'IS': # individual screen per pair or triad or quad
                            corenumber = CableDict[item][1][0]*CableDict[item][1][1] # total number of cores
                            multiplier = CableDict[item][1][1]
                            counter = 0
                            screennumber = 0
                            for i in range(1,corenumber+1):                       
                                LSignal = Name + "LSignal"+str(i)
                                #print(LSignal)
                                LTermination = Name + "LTermination" + str(i)                      
                                Color = Name + "Color" + str(i)                       
                                RTermination = Name + "RTermination" + str(i)
                                RSignal = Name + "RSignal"+str(i)
                                if condition1:
                                    Terminations.ActiveSheet.Cells(row,1).Value = CableDict[item][i+2]    # LSignal starts on index 3
                                    Terminations.ActiveSheet.Cells(row,1).Name = LSignal
                                if condition2:
                                    Terminations.ActiveSheet.Cells(row,7).Value = CableDict[item][i+302]  # RSignal starts on index 303
                                    Terminations.ActiveSheet.Cells(row,7).Name = RSignal
                                if condition1:
                                    Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,2).Name = LTermination
                                    Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,6).Name = RTermination
                                if condition2:
                                   Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"
                                   Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                   Terminations.ActiveSheet.Cells(row,6).Name = LTermination
                                   Terminations.ActiveSheet.Cells(row,6).Interior.Color = rgbToInt((226,239,218))
                                   Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                                   Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                   Terminations.ActiveSheet.Cells(row,2).Name = RTermination
                                Terminations.ActiveSheet.Cells(row,3).Value = CableDict[item][i+352] #  CoreNumber starts at index 353 
                                Terminations.ActiveSheet.Cells(row,4).Value = CableDict[item][i+152]  # Color start at index 153
#                                Terminations.ActiveSheet.Cells(row,4).Name = Color
                                Terminations.ActiveSheet.Cells(row,5).Value = CableDict[item][i+352] # CoreNumber starts at index 353
                                row += 1
                                counter += 1
                                if counter == multiplier:
                                    screennumber += 1
                                    LSCR = Name + "LSCR" + str(screennumber)
                                    RSCR = Name + "RSCR" + str(screennumber)
                                    if condition1:
                                        Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][screennumber+102]   # LSCR starts on index 103
                                        Terminations.ActiveSheet.Cells(row,2).Name = LSCR
                                        Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
                                        Terminations.ActiveSheet.Cells(row,6).Name = RSCR
                                    if condition2:
                                        Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][screennumber+102]   # LSCR starts on index 103
                                        Terminations.ActiveSheet.Cells(row,6).Name = LSCR                                        
                                        Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
                                        Terminations.ActiveSheet.Cells(row,2).Name = RSCR
                                    Terminations.ActiveSheet.Cells(row,3).Value = "SCR"+str(screennumber)                            
                                    Terminations.ActiveSheet.Cells(row,5).Value = "SCR"+str(screennumber)
                                    row +=1
                                    counter = 0 # reset the multiplier counter
                        if CableDict[item][1][2] == 'OS' or CableDict[item][1][2] == 'PE':
                            screenconfig = CableDict[item][1][2]
                            corenumber = CableDict[item][1][0]*CableDict[item][1][1] # total number of cores
                            screennumber = 1
                            for i in range(1,corenumber+1):                       
                                LSignal = Name + "LSignal"+str(i)
                                #print(LSignal)
                                LTermination = Name + "LTermination" + str(i)                      
                                Color = Name + "Color" + str(i)                       
                                RTermination = Name + "RTermination" + str(i)
                                RSignal = Name + "RSignal"+str(i)
                                Terminations.ActiveSheet.Cells(row,1).Value = CableDict[item][i+2]    # LSignal starts on index 3
                                Terminations.ActiveSheet.Cells(row,1).Name = LSignal
                                if condition1:
                                    Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,2).Name = LTermination
                                    Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,6).Name = RTermination 
                                if condition2:
                                    Terminations.ActiveSheet.Range("F"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][i+52]   # LTermination starts on index 53
                                    Terminations.ActiveSheet.Cells(row,6).Name = LTermination
                                    Terminations.ActiveSheet.Range("B"+str(row)).NumberFormat = "@"
                                    Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][i+252]  # RTermination starts on index 253
                                    Terminations.ActiveSheet.Cells(row,2).Name = RTermination                                     
                                Terminations.ActiveSheet.Cells(row,3).Value = CableDict[item][i+352] #  CoreNumber starts at index 353 
                                Terminations.ActiveSheet.Cells(row,4).Value = CableDict[item][i+152]  # Color start at index 153
                                Terminations.ActiveSheet.Cells(row,4).Name = Color
                                Terminations.ActiveSheet.Cells(row,5).Value = CableDict[item][i+352] #  CoreNumber starts at index 353
                                Terminations.ActiveSheet.Cells(row,7).Value = CableDict[item][i+302]  # RSignal starts on index 303
                                Terminations.ActiveSheet.Cells(row,7).Name = RSignal                            
                                row += 1
                            LSCR = Name + "LSCR" + str(screennumber)
                            RSCR = Name + "RSCR" + str(screennumber)
                            Terminations.ActiveSheet.Cells(row,2).Value = CableDict[item][screennumber+102]   # LSCR starts on index 103
                            Terminations.ActiveSheet.Cells(row,2).Name = LSCR                    
                            Terminations.ActiveSheet.Cells(row,3).Value = screenconfig                            
                            Terminations.ActiveSheet.Cells(row,5).Value = screenconfig
                            Terminations.ActiveSheet.Cells(row,6).Value = CableDict[item][screennumber+202]  # RSCR starts on index 203
                            Terminations.ActiveSheet.Cells(row,6).Name = RSCR                           
                            row +=1
                    if type(CableDict[item][1]) is str:
                        Terminations.ActiveSheet.Cells(row,4).Value = CableDict[item][1]
                        row += 1
                    HandledCables.append(item)
                    Terminations.ActiveSheet.Range("H"+str(startrow2)+":"+"H"+str(row-1)).Merge()
                    Terminations.ActiveSheet.Range("H"+str(startrow2)+":"+"H"+str(row-1)).VerticalAlignment = -4160 # Align top
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(1).LineStyle = 1 # Continous line
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(2).LineStyle = 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(1).Weight = 4 # Thick linestyle
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(2).Weight = 4
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(4).LineStyle = 1 # internal cell borders
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(3).LineStyle = 1
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(4).Weight = 2 # internal cell borders Thin
            Terminations.ActiveSheet.Range("A"+str(startrow)+":"+"H"+str(row-1)).Borders(3).Weight = 2
            Terminations.ActiveSheet.Range("A:H").Columns.AutoFit()
            Terminations.ActiveSheet.Range("A:H").HorizontalAlignment = xlHAlignCenter
            xlUp = -4162
            LastRow = Terminations.ActiveSheet.Cells(Terminations.ActiveSheet.Rows.Count, "C").End(xlUp).Row
            Terminations.ActiveSheet.Range("E2:G"+str(LastRow)).Interior.Color = rgbToInt((226,239,218))
        time.sleep(2)
        Terminations.Worksheets("Sheet1").Delete()
        Terminations.Close(SaveChanges=True)
        Xcel.Application.Quit()
        tk.messagebox.showwarning(title=None, message="Enclosures Terminations finished.")
        btn_Terminations.config(state=tk.NORMAL)                 
    except BaseException as e:
        print(e.args)
        btn_Terminations.config(state=tk.NORMAL) 
        connection.close()
        storage.close()        
        progress.stop()
          
def cleanup():
    print("CLEANUP")
    global Datasheet, Object
    window.destroy()
    gc.collect() # apparently to remove any threads
    os.system('TASKKILL /F /IM excel.exe')
    
def test():
    for thread in threading.enumerate():
        print(thread.name)
   
#window = tk.Tk()
#window.title("EDB")
window = tk.Tk()
window.title("PinDB Dashboard")
#window.configure(background='gray40')

window.rowconfigure([0, 1, 2, 3, 4, 5, 6, 7, 8, 9], minsize=30, weight=1)
window.columnconfigure([0, 1, 2, 3, 4, 5, 6], weight=1)

#VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV GUI Section VVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVVV
#================== SECTION DIAGRAM CRAWLER ===============================================
lbl_crawl = tk.Label(master=window,text="DIAGRAM Examination.",justify=tk.LEFT,relief=tk.GROOVE)
lbl_crawl.grid(row=0, column=0, sticky="nsew", padx=10)

# Vertical scrollbar for lbox_tags below
scrb_tags = tk.Scrollbar(orient=tk.VERTICAL, troughcolor = 'gray40')
scrb_tags.grid(row=1, column=1, rowspan=2, sticky="nsw")

# Object selection via listbox
lbox_tags = tk.Listbox(master=window, selectmode=tk.SINGLE, height=18, width=30, yscrollcommand=scrb_tags)
lbox_tags.bind('<Double-1>',choosetag)
lbox_tags.grid(row=1, rowspan=2, column=0, sticky="nsew", padx=10)
scrb_tags['command'] = lbox_tags.yview

btn_diagram = tk.Button(master=window, text="SELECT DIAGRAM", command=select)
btn_diagram.grid(row=3, column=0, sticky="nsew", padx=10)
btn_diagram_ttp = CreateToolTip(btn_diagram,\
                                "Select a diagram from the fileselector. "
                                "Then click on FIND TAGS to find any tags "
                                "in the diagram.")

lbl_diagram = tk.Label(master=window,text="Selected Diagram",justify=tk.LEFT,relief=tk.GROOVE)
lbl_diagram.grid(row=4, column=0, sticky="nsew", padx=10)

btn_crawl = tk.Button(master=window, text="FIND TAGS", command=startcrawl)
btn_crawl.grid(row=5, column=0, sticky="nsew", padx=10)

btn_tagexport = tk.Button(master=window, text="EXPORT TAGS", command=tagexport)
btn_tagexport.grid(row=6, column=0, sticky="nsew", padx=10)

btn_connections = tk.Button(master=window, text="FIND CONNECTIONS", pady=4, command=startconnections)
btn_connections.grid(row=7, column=0, sticky="nsew", padx=10)

btn_loopexport = tk.Button(master=window, text="EXPORT CONNECTIONS", pady=4, command=loopexport)
btn_loopexport.grid(row=8, column=0, sticky="nsew", padx=10)

# Progress Bar to show progress of P&ID walkthrough
progress = ttk.Progressbar(mode='indeterminate')
progress.grid(row=9, column=0, sticky="nsew", padx=10)

#================== SECTION DATASHEETS  ===============================================
lbl_new = tk.Label(master=window,text="Choose a Datasheet\nfrom the list below\nand click NEW OBJECT.",justify=tk.LEFT,relief=tk.GROOVE)
lbl_new.grid(row=0, column=2, sticky="nsew", padx=10)

# Vertical scrollbar for lbox_datasheets below
scrb_datasheets = tk.Scrollbar(orient=tk.VERTICAL)
scrb_datasheets.grid(row=1, column=3, rowspan=2, sticky="nsw")
# Datasheet selection via Listbox
lbox_datasheets = tk.Listbox(master=window, selectmode=tk.SINGLE, height=18, width=30, yscrollcommand=scrb_datasheets)
for item in DS_List:
    lbox_datasheets.insert(tk.END, item)
lbox_datasheets.grid(row=1, rowspan=2, column=2, sticky="nsew", padx=10)
scrb_datasheets['command'] = lbox_datasheets.yview

btn_new = tk.Button(master=window, text="NEW OBJECT", command=lambda: new("NOTAG"))
btn_new.grid(row=3, column=2, sticky="nsew", padx=10)

btn_commit_new = tk.Button(master=window, text="COMMIT NEW OBJECT", command=commit)
btn_commit_new.grid(row=4, column=2, sticky="nsew", padx=10)

lbl_chosenTag = tk.Label(master=window,text="No chosen Tag",justify=tk.LEFT,relief=tk.GROOVE)
lbl_chosenTag.grid(row=5, column=2, sticky="nsew", padx=10)

btn_bulkimport = tk.Button(master=window, text="BULK TAG IMPORT", command=bulkimport)
btn_bulkimport.grid(row=6, column=2, sticky="nsew", padx=10)

btn_loopimport = tk.Button(master=window, text="CONNECTIONS IMPORT", command=loopimport)
btn_loopimport.grid(row=7, column=2, sticky="nsew", padx=10)

#================== SECTION DATABASE OBJECTS ===============================================
lbl_retain = tk.Label(master=window,text="Choose an object\nfrom the list below\nand click RETAIN\nREMOVE,RENAME or COPY.",justify=tk.LEFT,relief=tk.GROOVE)
lbl_retain.grid(row=0, column=4, sticky="nsew", padx=10)

# Vertical scrollbar for lbox_objects below
scrb_objects = tk.Scrollbar(orient=tk.VERTICAL)
scrb_objects.grid(row=1, column=5, rowspan=2, sticky="nsw")
# Object selection via listbox
lbox_objects = tk.Listbox(master=window, selectmode=tk.EXTENDED, height=18, width=30, yscrollcommand=scrb_objects)
OBJ_List.sort()
for item in OBJ_List:
    lbox_objects.insert(tk.END, item)
lbox_objects.grid(row=1, rowspan=2, column=4, sticky="nsew", padx=10)
lbox_objects.bind('<Button-3>',fillobject)
lbox_objects.bind('<Double-Button-1>',showclass)
scrb_objects['command'] = lbox_objects.yview
lbox_objects_ttp = CreateToolTip(lbox_objects,\
                                 "If a category of objects is chosen "
                                 "(via OBJECT TYPE selection)"
                                 "and objects are selected in this list, "                              
                                 "right click and a list with prefilled "
                                 "datasheets will be opened to transfer "
                                 "to the objects.")

btn_retain = tk.Button(master=window, text="RETAIN", command=retain)
btn_retain.grid(row=3, column=4, sticky="nsew", padx=10)

btn_commit = tk.Button(master=window, text="COMMIT OBJECT", command=commit)
btn_commit.grid(row=4, column=4, sticky="nsew", padx=10)

btn_remove = tk.Button(master=window, text="REMOVE", command=remove)
btn_remove.grid(row=5, column=4, sticky="nsew", padx=10)

btn_rename = tk.Button(master=window, text="RENAME", command=rename)
btn_rename.grid(row=6, column=4, sticky="nsew", padx=10)

btn_copy = tk.Button(master=window, text="COPY", command=copyobject)
btn_copy.grid(row=7, column=4, sticky="nsew", padx=10)

btn_reclass = tk.Button(master=window, text="RECLASSIFY", command=reclassify)
btn_reclass.grid(row=8, column=4, sticky="nsew", padx=10)

lbl_objectnumber = tk.Label(master=window,text="No of Objects:",justify=tk.LEFT,relief=tk.GROOVE)
lbl_objectnumber.grid(row=9, column=4, sticky="nsew", padx=10)
lbl_objectnumber["text"] = "No of Objects: "+str(len(root))

#================== SECTION DATABASE BULK FUNCTIONS  ===============================================
# These are grouped in a Frame for neatness

lbl_functions = tk.Label(master=window,text="Database Operations",width=20,justify=tk.LEFT,relief=tk.GROOVE)
lbl_functions.grid(row=0, column=6, sticky="nsew", padx=10)

frm_dbfuncs =  tk.Frame(master=window)
frm_dbfuncs.grid(row=1, rowspan=2, column=6, padx=10)

lbl_documents = tk.Label(master=frm_dbfuncs,text="Databas Export/Import",width=25,height=3,justify=tk.LEFT,relief=tk.GROOVE)
lbl_documents.grid(row=0, column=0, sticky="nsew", padx=10)

btn_index = tk.Button(master=frm_dbfuncs, text="INSTRUMENT INDEX", command=index)
btn_index.grid(row=1, column=0, sticky="ew", padx=10)

btn_Inimport = tk.Button(master=frm_dbfuncs, text="IMPORT INSTR-INDEX", command=startindeximport)
btn_Inimport.grid(row=2, column=0, sticky="ew", padx=10)

btn_terminationexport = tk.Button(master=frm_dbfuncs, text="EXPORT TERMINATIONS", command=terminationexport)
btn_terminationexport.grid(row=3, column=0, sticky="ew", padx=10)

btn_terminationimport = tk.Button(master=frm_dbfuncs, text="IMPORT TERMINATIONS", command=terminationimport)
btn_terminationimport.grid(row=4, column=0, sticky="ew", padx=10)

btn_objectmenu = tk.Menubutton(master=frm_dbfuncs, text="OBJECT TYPE",relief=tk.GROOVE,bg="green",fg="white")
btn_objectmenu.grid(row=5, column=0, sticky="ew", padx=10)

ObjectChoice = tk.StringVar()

menu_objectmenu = tk.Menu(btn_objectmenu)
btn_objectmenu["menu"] = menu_objectmenu
menu_objectmenu.add_radiobutton(label = "ALL", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "ANALOG INPUTS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "ANALOG OUTPUTS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "ANTENNAS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "BEACONS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "CABLES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "COMPASSES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "CONTROL VALVES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "DIGITAL INPUTS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "DIGITAL OUTPUTS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "ENCLOSURES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "FIRE DETECTORS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "FLOWMETERS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "FOGDETECTORS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "FOGHORNS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "GAS DETECTORS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "HANDSWITCHES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "LEVEL GAUGES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "LIMIT SWITCHES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "LEVEL TRANSMITTERS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "LOADPINS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "OCEANOGRAPHS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "PRESSURE GAUGES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "PIG DETECTORS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "PRESSURE TRANSMITERS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "RESTRICTION ORIFICES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "SLIPRINGS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "SPEAKERS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "SOLENOIDS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "SOLARPANELS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "SAFETY VALVES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "TEMPERATURE GAUGES", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "TEMPERATURE TRANSMITTERS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "WEATHERSTATIONS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "WINDGENERATORS", variable=ObjectChoice, command=objects)
menu_objectmenu.add_radiobutton(label = "TRANSFORMERS", variable=ObjectChoice, command=objects)

lbl_objectchoice = tk.Label(master=frm_dbfuncs, text="ALL", width=25, justify=tk.LEFT,relief=tk.GROOVE)
lbl_objectchoice.grid(row=6, column=0, sticky="nsew", padx=10)

btn_export = tk.Button(master=frm_dbfuncs, text="EXPORT SELECTION", command=startexport)
btn_export.grid(row=7, column=0, sticky="ew", padx=10)

btn_import = tk.Button(master=frm_dbfuncs, text="IMPORT SELECTION", command=startimport)
btn_import.grid(row=8, column=0, sticky="ew", padx=10)

btn_Autocad = tk.Button(master=frm_dbfuncs, text="AUTOCAD EXPORT", command=insertAcad)
btn_Autocad.grid(row=9, column=0, sticky="ew", padx=10)

btn_objectnumbers = tk.Button(master=frm_dbfuncs, text="OBJECT NUMBERS", command=objectnumbers)
btn_objectnumbers.grid(row=10, column=0, sticky="ew", padx=10)

btn_PCexport = tk.Button(master=frm_dbfuncs, text="PC EXPORT", command=startPCexport)
btn_PCexport.grid(row=11, column=0, sticky="ew", padx=10)

btn_PCimport = tk.Button(master=frm_dbfuncs, text="PC IMPORT", command=startPCimport)
btn_PCimport.grid(row=12, column=0, sticky="ew", padx=10)

lbl_documents = tk.Label(master=frm_dbfuncs,text="Document Export Functions",width=25,height=3,justify=tk.LEFT,relief=tk.GROOVE)
lbl_documents.grid(row=13, column=0, sticky="nsew", padx=10)

btn_collect = tk.Button(master=frm_dbfuncs, text="COLLECT DATASHEETS", command=collect)
btn_collect.grid(row=14, column=0, sticky="ew", padx=10)

btn_Cableschedule = tk.Button(master=frm_dbfuncs, text="CABLE SCHEDULE", command=cableschedule)
btn_Cableschedule.grid(row=15, column=0, sticky="ew", padx=10)

btn_IOlist = tk.Button(master=frm_dbfuncs, text="I/O LIST", command=iolist)
btn_IOlist.grid(row=16, column=0, sticky="ew", padx=10)

btn_Terminations = tk.Button(master=frm_dbfuncs, text="ENCLOSURE TERMINATIONS", command=enclosureterminations)
btn_Terminations.grid(row=17, column=0, sticky="ew", padx=10)

#btn_test = tk.Button(master=frm_dbfuncs, text="TEST", command=test)
#btn_test.grid(row=18, column=0, sticky="ew", padx=10)

#btn_cleanup = tk.Button(master=window, text="CLEANUP", command=cleanup)
#btn_cleanup.grid(row=1, column=1, sticky="nsew")

window.protocol("WM_DELETE_WINDOW", cleanup)
window.mainloop()