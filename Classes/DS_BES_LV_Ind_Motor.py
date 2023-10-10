# -*- coding: utf-8 -*-

class DS_BES_Transformer():
    TagNumber = ""
    AssetNumber = ""
    Specification = ""
    AssociatedEquipment = ""
    PandID = ""
    ServiceDescription = ""
    Manufacturer = ""
    Model = ""
    PO_Number = ""
    REQ_Number = ""
    PlaceManufacturer = ""
    LocalServiceOrg = ""
    Revision = ""
    Originator = ""
    Checker = ""
    Approver = ""
    Date = ""
    RevDescription = ""
    Project = ""
    SheetNumber = ""
    ProjectNumber = ""
    BESDrawingNumber = ""
    ClientDrawingNumber = ""
    VesselNumber = ""
    LineID = ""
    Size = ""
    Schedule = ""
    FieldsDict = {
    "TagNumber": "_TAG_NO",
    "AssetNumber": "_Asset_No",
    "Specification": "_SpecNo",
    "Associated Equipment": "__ASSOC_EQ",
    "PandID": "_PID_NO",
    "ServiceDescription": "_Service",
    "Manufacturer": "_Manufactr",
    "Model": "_Model",
    "PO_Number": "_P_ORDER",
    "REQ_Number": "_REQ_NO",
    "PlaceManufacturer": "_PlOfManufac",
    "LocalServiceOrg": "_LoServOrga",
    "Revision": "Rev",
    "Originator": "RevBy",
    "Checker": "Chk",
    "Approver": "App",
    "Date": "RevDate",
    "RevDescription": "RevStatus",
    "Project": "_Projectname",
    "SheetNumber": "_JSHEET_NO",
    "ProjectNumber": "_Projectnumber",
    "BESDrawingNumber": "_DOC_NAME",
    "ClientDrawingNumber": "_ClntDwg",
    "VesselNumber": "_VesselNo",
    "LineID": "_LINE_NO",
    "Size": "_Size",
    "Schedule": "_Schedule"
    }   
    InternalFieldsDict = {
        "AFHVRatedCurrent": "AFHVRatedCurrent",}