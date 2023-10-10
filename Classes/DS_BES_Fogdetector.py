# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Fogdetector(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    # DETECTOR
    Power = ""
    Height = ""
    PowerConsumption = ""
    Length = ""
    WarmupTime = ""
    Width = ""
    UpdateTime = ""
    Mounting = ""
    Wavelength = ""
    Gland = ""
    VisibilityRange = ""
    Approval = ""
    Outputs = ""
    HousingMaterial = ""
    Connection1 = ""
    Connection2 = ""     
    ElectricalProtection = ""
    GasGroup = ""
    TempCategory = ""
    IPRating = ""
    # ACCESSORY
    MountingKit = ""
    Other_Field1 = ""
    Other_Field2 = ""
    Other_Field3 = ""
    Other_Field4 = ""
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
    "ActivationSource": "_GenActivSrc",
    "Function": "_GenFunct",
    "Power": "_DetectorPwrReqs",
    "Height": "_DetectorHght",
    "Power Consumption": "_DetectorPwrConsumption",
    "Length": "_DetectorLength",
    "WarmupTime": "_DetectorWarmupTime",
    "Width": "_DetectorWidth",
    "UpdateTime": "_DetectorUpdateTime",
    "Mounting": "_DetectorMount",
    "Wavelength": "_DetectorWavelength",
    "Gland": "_DetectorGlandConn",
    "VisibilityRange": "_DetectorVisibilityRange",
    "Approval": "_DetectorApproval",
    "Outputs": "_DetectorOutputs",
    "HousingMaterial": "_DetectorEnclMat",
    "Connection1":"_Connection1",
    "Connection2":"_Connection2",    
    "ElectricalProtection": "_PROT_TYPE",
    "GasGroup": "_GASGROUP",
    "TempCategory": "_TEMP_CLASS",
    "IPRating": "_IP_RATING",
    "MountingKit": "_AccMountingKit",
    "Other_Field1": "_AccOther1",
    "Other_Field2": "_AccOther2",
    "Other_Field3": "_AccOther3",
    "Other_Field4": "_AccOther4",
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

    