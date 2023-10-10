# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Solarpanel(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    Function = ""
    # ELECTRICAL    
    NominalVoltage = ""
    RatedPower = ""
    OpenCircuitVoltage = ""
    MaxPowerVoltage = ""
    MaxCurrent = ""
    ShortCircuitCurrent = ""
    Efficiency = ""
    MaximumPower = ""
    OutputTolerance = ""
    Diode = ""
    MonoPolycrystalline = ""
    TempCoefficientofVoc = ""
    TempCoefficientofIsc = ""
    TempCoefficientofPmax = ""
    Connection1 = ""
    Connection2 = ""
    # CONSTRUCTION
    Connector = ""
    FrameMaterial = ""
    FrameType = ""
    Locking = ""
    GlassCover = ""
    Length = ""
    Width = ""
    Thickness = ""
    Weight = ""
    Mounting = ""
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
    "NominalVoltage": "_ElecNominalVoltage",
    "RatedPower": "_ElecRatedPower",
    "OpenCircuitVoltage": "_ElecOpenCircuitVoltage",
    "MaxPowerVoltage": "_ElecMaxPowerVoltage",
    "MaxCurrent": "_ElecMaxCurrent",
    "ShortCircuitCurrent": "_ElecShortCircuitCurrent",
    "Efficiency": "_ElecEfficiency",
    "MaximumPower": "_ElecMaximumPower",
    "OutputTolerance": "_ElecOutputTolerance",
    "Diode": "_ElecDiode",
    "MonoPolycrystalline": "_ElecMonoPolyCrystalline",
    "TempCoefficientofVoc": "_ElecTCVoc",
    "TempCoefficientofIsc": "_ElecTCIsc",
    "TempCoefficientofPmax": "_ElecTCPmax",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",
    "Connector": "_CnstrConnector",
    "FrameMaterial": "_CnstrFrameMaterial",
    "FrameType": "_CnstrFrameType",
    "Locking": "_CnstrLocking",
    "GlassCover": "_CnstrGlassCover",
    "Length": "_CnstrLength",
    "Width": "_CnstrWidth",
    "Thickness": "_CnstrThickness",
    "Weight": "_CnstrWeigth",
    "Mounting": "_CnstrMounting",
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

    