# -*- coding: utf-8 -*-
#from Classes.DS_BES_Beacon_atex import DS_BES_Beacon_atex
#from Classes.DS_BES_Cable import DS_BES_Cable
#from Classes.DS_BES_CV_atex import DS_BES_CV_atex
#from Classes.DS_BES_FD_atex import DS_BES_FD_atex
#from Classes.DS_BES_FI_atex import DS_BES_FI_atex
#from Classes.DS_BES_GD_atex import DS_BES_GD_atex
#from Classes.DS_BES_Handsw_atex import DS_BES_Handsw_atex
#from Classes.DS_BES_LG import DS_BES_LG
#from Classes.DS_BES_Limitsw_atex import DS_BES_Limitsw_atex
#from Classes.DS_BES_LIT_atex import DS_BES_LIT_atex
#from Classes.DS_BES_PG import DS_BES_PG
#from Classes.DS_BES_PIT_atex import DS_BES_PIT_atex
#from Classes.DS_BES_RO import DS_BES_RO
#from Classes.DS_BES_SOL_V_atex import DS_BES_SOL_V_atex
#from Classes.DS_BES_SV import DS_BES_SV
#from Classes.DS_BES_TG import DS_BES_TG
#from Classes.DS_BES_TT_atex import DS_BES_TT_atex
# import ZODB and supporting libraries
from ZODB import FileStorage, DB
import transaction
import win32com.client, os, time

#++++ PACKING OF DATABASE
#'''
storage = FileStorage.FileStorage('Databases\Benin.fs')
db = DB(storage)
db.pack()
storage.close()
#++++ OBTAING DATABASE CONTENT
#connection = db.open()
#root = connection.root()
#for key in root:
#    Object = root[key]
#    print(key,isinstance(Object,DS_BES_Cable))
#print(root)
#del root['']
#del root[None]
#transaction.commit()
#connection.close()
#storage.close()
#'''
##++++ OPEN ACAD DRAWING
#'''
#directory = os.getcwd()+"/ACAD"
#filename = "/RAS-B-840-DP-1002-001-3.DWG"
#acad = win32com.client.Dispatch("Autocad.Application")
#doc = acad.Documents.Open(directory+filename)
#acad.Visible = False
#for entity in acad.ActiveDocument.ModelSpace:
#    name = entity.ObjectName 
#    if name == 'AcDbBlockReference':
##            print(entity.InsertionPoint)
#            HasAttributes = entity.HasAttributes
#            if HasAttributes:
#                for attrib in entity.GetAttributes():
#                    print(attrib.TagString,attrib.TextString)
#    time.sleep(0.3)                    
#doc.close()
#os.system('TASKKILL /F /IM acad.exe')