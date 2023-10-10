# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_PIT(Instrument):
    # PROCESS CONDITIONS
    HiPressureFluid = ""
    HiGravityAtOpTemp = ""
    HiViscosityAtOpTemp = ""
    HiMinPressure = ""
    HiNormPressure = ""
    HiMaxPressure = ""
    HiMinTemperature = ""
    HiNormTemperature = ""
    HiMaxTemperature = ""
    HISolids = ""
    HiQuality = ""
    HiService = ""
    HiCritical = ""
    HiPulsating = ""
    AreaClassification = ""
    LoPressureFluid = ""
    LoGravityAtOpTemp = ""
    LoViscosityAtOpTemp = ""
    LoMinPressure = ""
    LoNormPressure = ""
    LoMaxPressure = ""
    LoMinTemperature = ""
    LoNormTemperature = ""
    LoMaxTemperature = ""
    LoSolids = ""
    LoQuality = ""
    LoService = ""
    LoCritical = ""
    LoPulsating = ""
    AmbTemperatureReq = ""
    SpecificGravityUnits = ""
    ViscosityUnits = ""
    PressUnits = ""
    TempUnits = ""
    # ELEMENT
    ElementType = ""
    ElementFillFluid = ""
    ElementMinSpan = ""
    ElementMaxSpan = ""
    ElementDiaphragmMat = ""
    ElementDrainLocation = ""
    ElementDrainMat = ""
    ElementProcConn = ""
    # TRANSMITTER
    TransmitterVoltage = ""
    TransmitterPowerWiring = ""
    TransmitterSignalType = ""
    TransmitterProtocol = ""
    TransmitterLocation = ""
    TransmitterSmart = ""
    TransmitterIndicate = ""
    TranmitterIsolate = ""
    TransmitterElectricalProtection = ""
    TransmitterTempCategory = ""
    TransmitterGasGroup = ""
    TransmitterIPRating = ""
    TransmitterFacCalibration = ""
    TransmitterTempComp = ""
    TransmitterCharacteristic = ""
    TransmitterCalRange = ""
    TransmitterOverRange = ""
    TransmitterZeroElevation = ""
    TransmitterAccuracy = ""
    TransmitterGasketMat = ""
    TransmitterGlandConn = ""
    TransmitterNACE = ""
    TransmitterElemConn = ""
    TransmitterSSTag = ""
    TransmitterBodyMaxPress = ""
    TransmitterBodyMat = ""
    TransmitterMount = ""
    # SEAL DIAPHRAM & CAPILLARY
    SealDiaLength = ""
    SealDiaID = ""
    SealDiaArmor = ""
    SealDiaFillFluid = ""
    SealDiaMaxResponseTime = ""
    SealDiaSGat60oF = ""
    SealDiaCapMat = ""
    SealDiaHiPressSizeType = ""
    SealDialHiPressFlushRing = ""
    SealDiaLoPressSizeType = ""
    SealDiaLoPressThickness = ""
    SealDialLoPressMat = ""
    SealDialLoPressFlushRing = ""
    SealDiaTempRating = ""
    SealDiaMaxTemp = ""
    SealDIaPressRating = ""
    SealDiaMaxPress = ""
    SealDIaManufacturer = ""
    SealDiaModel = ""
    # MANIFOLD
    MnfldType = ""
    MnfldMat = ""
    MnfldTransmitterConn = ""
    MnfldProcConn = ""
    MnfldManufacturer = ""
    MnfldModel = ""
    # CONNECTIONS
    Connection1 = ""
    Connection2 = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Note_Field5 = ""
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
    InternalFieldsDict = {
    "HiPressureFluid": "_PCHiPressFluid",
    "HiGravityAtOpTemp": "_PCHiPressConnSpecificGravityAtOpTemp",
    "HiViscosityAtOpTemp": "_PCHiPressConnViscosityAtOpTemp",
    "HiMinPressure": "_PCHiMinPress",
    "HiNormPressure": "_PCHiNormPress",
    "HiMaxPressure": "_PCHiMaxPress",
    "HiMinTemperature": "_PCHiPressMinTemp",
    "HiNormTemperature": "_PCHiPressNormTemp",
    "HiMaxTemperature": "_PCHiPressMaxTemp",
    "HISolids": "_PCHiPressPercSolids",
    "HiQuality": "_PCHiPressPercQuality",
    "HiService": "_PCHiPressService",
    "HiCritical": "_PCHiPressCritical",
    "HiPulsating": "_PCHiPressPulsating",
    "AreaClassification": "_PCAREA_CLASS_REQ",
    "LoPressureFluid": "_PCLowPressFluid",
    "LoGravityAtOpTemp": "_PCLowPressConnSpecificGravityAtOpTemp",
    "LoViscosityAtOpTemp": "_PCLowPressConnViscosityAtOpTemp",
    "LoMinPressure": "_PCLowMinPress",
    "LoNormPressure": "_PCLowNormPress",
    "LoMaxPressure": "_PCLowMaxPress",
    "LoMinTemperature": "_PCLowPressMinTemp",
    "LoNormTemperature": "_PCLowPressNormTemp",
    "LoMaxTemperature": "_PCLowPressMaxTemp",
    "LoSolids": "_PCLowPressPercSolids",
    "LoQuality": "_PCLowPressPercQuality",
    "LoService": "_PCLowPressService",
    "LoCritical": "_PCLowPressCritical",
    "LoPulsating": "_PCLowPressPulsating",
    "AmbTemperatureReq": "_PCAmbTempReqs",
    "SpecificGravityUnits": "_PCSpGravityUnits",
    "ViscosityUnits": "_PCViscosityUnits",
    "PressUnits": "_PCPressUnits",
    "TempUnits": "_PCTempUnits",
    "ElementType": "_ElmntType",
    "ElementFillFluid": "_ElmntFillFluid",
    "ElementMinSpan": "_ElmntMinSpan",
    "ElementMaxSpan": "_ElmntMaxSpan",
    "ElementDiaphragmMat": "_ElmntDiaphWetMat",
    "ElementDrainLocation": "_ElmntVentDrainLocation",
    "ElementDrainMat": "_ElmntVentDrainMat",
    "ElementProcConn": "_ElmntProcConn",
    "TransmitterVoltage": "_XmtrVolt",
    "TransmitterPowerWiring": "_XmtrPwrWiring",
    "TransmitterSignalType": "_XmtrSigType",
    "TransmitterProtocol": "_XmtrCommProtocol",
    "TransmitterLocation": "_XmtrLocation",
    "TransmitterSmart": "_XmtrSmart",
    "TransmitterIndicate": "_XmtrIndicate",
    "TranmitterIsolate": "_XmtrIsolate",
    "TransmitterElectricalProtection": "_PROT_TYPE",
    "TransmitterTempCategory": "_TEMP_CLASS",
    "TransmitterGasGroup": "_GASGROUP",
    "TransmitterIPRating": "_IP_RATING",
    "TransmitterFacCalibration": "_XmtrFactCalibration",
    "TransmitterTempComp": "_XmtrAmbTempCompens",
    "TransmitterCharacteristic": "_XmtrCharacteristic",
    "TransmitterCalRange": "_RANGE",
    "TransmitterOverRange": "_XmtrOvrRange",
    "TransmitterZeroElevation": "_XmtrZeroElevation",
    "TransmitterAccuracy": "_XmtrAccuracy",
    "TransmitterGasketMat": "_XmtrGaskMat",
    "TransmitterGlandConn": "_XmtrGlandConn",
    "TransmitterNACE": "_XmtrNACE",
    "TransmitterElemConn": "_XmtrElmntConn",
    "TransmitterSSTag": "_XmtrSSTag",
    "TransmitterBodyMaxPress": "_XmtrBdyMaxPressRating",
    "TransmitterBodyMat": "_XmtrBdyFlngMaterial",
    "TransmitterMount": "_XmtrMount",
    "SealDiaLength": "_SDCLen",
    "SealDiaID": "_SDCID",
    "SealDiaArmor": "_SDCArmor",
    "SealDiaFillFluid": "_SDCFillFluid",
    "SealDiaMaxResponseTime": "_SDCMaxRespTime",
    "SealDiaSGat60oF": "_SDCSGAt60F",
    "SealDiaCapMat": "_SDCCapMaterial",
    "SealDiaHiPressSizeType": "_SDCHiPressSizeAndType",
    "SealDialHiPressFlushRing": "_SDCHiPressFlushRing",
    "SealDiaLoPressSizeType": "_SDCLowPressSizeAndType",
    "SealDiaLoPressThickness": "_SDCLowPressThickness",
    "SealDialLoPressMat": "_SDCLowPressMaterial",
    "SealDialLoPressFlushRing": "_SDCLowPressFlushRing",
    "SealDiaTempRating": "_SDCTempRating",
    "SealDiaMaxTemp": "_SDCMaxTemp",
    "SealDIaPressRating": "_SDCPressRating",
    "SealDiaMaxPress": "_SDCMaxPress",
    "SealDIaManufacturer": "_SDCManufact",
    "SealDiaModel": "_SDCModel",
    "MnfldType": "_ManifoldType",
    "MnfldMat": "_ManifoldMaterial",
    "MnfldTransmitterConn": "_ManifoldXmtrConn",
    "MnfldProcConn": "_ManifoldProcConn",
    "MnfldManufacturer": "_ManifoldManufact",
    "MnfldModel": "_ManifoldModel",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Note_Field4": "_Notes4",
    "Note_Field5": "_Notes5",
    "Remarks": "_Remarks",
    "EC_TE_Certificate": "_EC_TYP_EX_CERT",
    "EC_DoC": "_EC_DECL_CONF",
    "QA_Notification": "_PROD_QA_NOTIFICATION",
    "QA_Not_Date": "_PROD_QA_NOT_DATE",
    "NOBO_Number": "_NOBO_REF",
    "User3": "_User3",
    "User4": "_User4",
    "User5": "_User5"
    }


