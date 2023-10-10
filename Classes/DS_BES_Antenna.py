# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Antenna(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    # ELECTRICAL    
    Model = ""
    Frequency = ""
    AntennaType = ""
    Polarisation = ""
    PatternType = ""
    ThreedBBeamwidth = ""
    Impedance = ""
    Connection1 = ""
    Gain = ""
    VSWR = ""
    MaxInputPower = ""
    Bandwidth = ""
    AntistaticProtection = ""
    HCMCodes = ""
    Connection1 = ""
    Connection2 = ""
    # MECHANICAL
    Connections = ""
    Materials = ""
    Colour = ""
    WindArea = ""
    WindLoad = ""
    DiameterAtTopEnd = ""
    DiameterAtBottomEnd = ""
    Height = ""
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
    "Model": "_ElecModel",
    "Frequency": "_ElecFrequency",
    "AntennaType": "_ElecAntennaType",
    "Polarisation": "_ElecPolarisation",
    "PatternType": "_ElecPatternType",
    "ThreedBBeamwidth": "_Elec3dBBeamwidth",
    "Impedance": "_ElecImpedance",
    "Connection1": "_Connection1",
    "Gain": "_ElecGain",
    "VSWR": "_ElecVSWR",
    "MaxInputPower": "_ElecMaxInputPower",
    "Bandwidth": "_ElecBandwidth",
    "AntistaticProtection": "_ElecAntistaticProtection",
    "HCMCodes": "_ElecHCMCodes",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "Connections": "_CnstrConnections",
    "Materials": "_CnstrMaterials",
    "Colour": "_CnstrColour",
    "WindArea": "_CnstrWindArea",
    "WindLoad": "_CnstrWindLoad",
    "DiameterAtTopEnd": "_CnstrDiameterTopEnd",
    "DiameterAtBottomEnd": "_CnstrDiameterBottomEnd",
    "Height": "_CnstrHeight",
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

    