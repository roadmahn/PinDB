# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Windgenerator(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    Function = ""
    # SPECIFICATIONS    
    RatedPower = ""
    VoltageSetPoints = ""
    VoltageRanges = ""
    FuseSizes = ""
    Kwhpermonth = ""
    TurbineController = ""
    Alternator = ""
    OverspeedProtection = ""
    StartupWindSpeed = ""
    MaximumWindSpeed = ""
    SurvivalWindSpeed = ""
    OperEnvironment = ""
    BladeMaterial = ""
    BodyMaterial = ""
    BladeHubMaterial = ""
    ShaftTrust = ""
    OperTempRange = ""
    StorageTempRange = ""
    RotorDiameter = ""
    Length = ""
    Width = ""
    SweptArea = ""
    Weight = ""
    Mounting = ""
    Connection1 = ""
    Connection2 = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Note_Field5 = ""
    Note_Field6 = ""
    Remarks = ""
    # ATEX  REGISTER
    EC_TE_Certificate = ""
    EC_DoC = ""
    QA_Notification = ""
    QA_Not_Date = ""
    NOBO_Number = ""
    User3 = ""
    User4 = ""
    User5 = ""
    User6 = ""
    User7 = ""
    User8 = ""
    User9 = ""
    InternalFieldsDict = {
    "AreaClassification": "_AREA_CLASS_REQ",
    "AmbientRelHumidity": "_GenAmbRelHumid",
    "AmbientTemperature": "_GenAmbTemp",
    "Function": "_GenFunct",
    "RatedPower": "_SpecRatedPower",
    "VoltageSetPoints": "_SpecVoltageSetPoints",
    "VoltageRanges": "_SpecVoltageRanges",
    "FuseSizes": "_SpecFuseSizes",
    "Kwhpermonth": "_SpecKwhperMonth",
    "TurbineController": "_SpecTurbineController",
    "Alternator": "_SpecAlternator",
    "OverspeedProtection": "_SpecOverspeedProtection",
    "StartupWindSpeed": "_SpecStartupWindSpeed",
    "MaximumWindSpeed": "_SpecMaxWindSpeed",
    "SurvivalWindSpeed": "_SpecSurvivalWindSpeed",
    "OperEnvironment": "_SpecOperatingEnvironment",
    "BladeMaterial": "_SpecBladeMaterial",
    "BodyMaterial": "_SpecBodyMaterial",
    "BladeHubMaterial": "_SpecBladeHubMaterial",
    "ShaftTrust": "_SpecShaftTrust",
    "OperTempRange": "_SpecOperatingTemp",
    "StorageTempRange": "_SpecStoreTemp",
    "RotorDiameter": "_SpecRotorDiameter",
    "Length": "_SpecLength",
    "Width": "_SpecWidth",
    "SweptArea": "_SpecSweptArea",
    "Weight": "_SpecWeight",
    "Mounting": "_SpecMounting",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",
    "Note_Field1": "_Note1",
    "Note_Field2": "_Note2",
    "Note_Field3": "_Note3",
    "Note_Field4": "_Note4",
    "Note_Field5": "_Note5",
    "Note_Field6": "_Note6",
    "Remarks": "_Remarks",    
    "EC_TE_Certificate": "_EC_TYP_EX_CERT",
    "EC_DoC": "_EC_DECL_CONF",
    "QA_Notification": "_PROD_QA_NOTIFICATION",
    "QA_Not_Date": "_PROD_QA_NOT_DATE",
    "NOBO_Number": "_NOBO_REF",
    "User3": "_User3",
    "User4": "_User4",
    "User5": "_User5",
    "User6": "_User6",
    "User7": "_User7",
    "User8": "_User8",
    "User9": "_User9"
    }

    