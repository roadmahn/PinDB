# -*- coding: utf-8 -*-
class ElecEquipment:
    TagNumber = ""
    PandID = ""
    ServiceDescription = ""
    PO_Number = ""
    Manufacturer = ""
    Model = ""
    Revision = ""
    Date = ""
    RevDescription = ""
    Originator = ""
    FieldsDict = {
    "TagNumber": "_TAG_NO",
    "PandID": "_PID_NO",
    "ServiceDescription": "_service",
    "PO_Number": "_P_ORDER",
    "Manufacturer": "_Manufactr",
    "Model": "_Model",
    "Revision": "Rev",
    "Date": "RevDate",
    "RevDescription": "RevStatus",
    "Originator": "RevBy"
    }