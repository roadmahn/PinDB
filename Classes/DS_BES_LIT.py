# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_LIT(Instrument):
    # PROCESS CONDITIONS = ""
    LowerFluid = ""
    LowerFluidDensityAtOpTemp = ""
    LowerFluidViscosityAtOpTemp = ""
    LowerFluidSolidsPercentage = ""
    LowerFluidDielecConst = ""
    LowerFluidService = ""
    LowerFluidCritical = ""
    UpperFluid = ""
    UpperFluidDensityAtOpTemp = ""
    UpperFluidViscosityAtOpTemp = ""
    UpperFluidSolidsPercentage = ""
    UpperFluidDielecConst = ""
    UpperFluidService = ""
    UpperFluidCritical = ""
    DensityUnits = ""
    ViscosityUnits = ""
    TempUnits = ""
    PressUnits = ""
    MinTemperature = ""
    NormTemperature = ""
    MaxTemperature = ""
    MinPressure = ""
    NormPressure = ""
    MaxPressure = ""
    AreaClassification = ""
    AmbTemperatureReq = ""
    # BODY / CAGE = ""
    CageType = ""
    CageUpperConnSize = ""
    CageUpperConnType = ""
    CageLowerConnSize = ""
    CageLowerConnType = ""
    CageMaterial = ""
    CageRating = ""
    CageRotatableHead = ""
    CageOrientation = ""
    CageDrainValve = ""
    CageVentValve = ""
    CageLength = ""
    CageDiameter = ""
    CageNozzleExtLength = ""
    CageNozzleSpan = ""
    CageCoolingExtension = ""
    CageMounting = ""
    # TRANSMITTER = ""
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
    TransmitterAccuracy = ""
    TransmitterSensorConn = ""
    TransmitterGlandConn = ""
    TransmitterBodyMaterial = ""
    TransmitterCalRange = ""
    TransmitterFullRange = ""
    TransmitterGasketMat = ""
    TransmitterSSTag = ""
    Connection1 = ""
    Connection2 = ""
    # ELEMENT = ""
    ElementType = ""
    ElementGauge = ""
    ElementGaugeMat = ""
    ElementGaugeDim = ""
    ElementInsertDepth = ""
    ElementProcConn = ""
    ElementTransmitterConn = ""
    ElementDim = ""
    ElementMat = ""
    ElementRadarFrequency = ""
    ElementDisplacerExt = ""
    ElementSpringTubeMat = ""
    ElementFloatWellClearance = ""
    ElementFloatShaftClearance = ""
    # SWITCHES = ""
    SwtchRelay1TagNumber = ""
    SwtchRelay1PowerRating = ""
    SwtchRelay1Type = ""
    SwtchRelay2PowerRating = ""
    SwtchRelay2Type = ""
    SwtchRelay2AlmSetting = ""
    SwtchCtctRating = ""
    SwtchGlandConn = ""
    SwtchManufacturer = ""
    SwtchModelNo = ""
    SwtchElectricalProtection = ""
    SwtchTempCat = ""
    SwtchGasGroup = ""
    SwtchIP1 = ""
    SwtchIP2 = ""
    # NOTES = ""
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Remarks = ""
    # ATEX  REGISTER = ""
    EC_TE_Certificate = ""
    EC_DoC = ""
    QA_Notification = ""
    QA_Not_Date = ""
    NOBO_Number = ""
    User3 = ""
    User4 = ""
    User5 = ""
    InternalFieldsDict = {
    "LowerFluid": "_PCLowerFluid",
    "LowerFluidDensityAtOpTemp": "_PCLowerFluidDensityAtOpTemp",
    "LowerFluidViscosityAtOpTemp": "_PCLowerFluidViscosityAtOpTemp",
    "LowerFluidSolidsPercentage": "_PCLowerFluidPercentSolids",
    "LowerFluidDielecConst": "_PCLowerFluidDielectricConst",
    "LowerFluidService": "_PCLowerFluidService",
    "LowerFluidCritical": "_PCLowerFluidCritical",
    "UpperFluid": "_PCUpperFluid",
    "UpperFluidDensityAtOpTemp": "_PCUpperFluidDensityAtOpTemp",
    "UpperFluidViscosityAtOpTemp": "_PCUpperFluidViscosityAtOpTemp",
    "UpperFluidSolidsPercentage": "_PCUpperFluidPercentSolids",
    "UpperFluidDielecConst": "_PCUpperFluidDielectricConst",
    "UpperFluidService": "_PCUpperFluidService",
    "UpperFluidCritical": "_PCUpperFluidCritical",
    "DensityUnits": "_PCSpGravityAtOpTempUnits",
    "ViscosityUnits": "_PCViscosityAtOpTempUnits",
    "TempUnits": "_PCTempUnits",
    "PressUnits": "_PCPressUnits",
    "MinTemperature": "_PCTempMin",
    "NormTemperature": "_PCTempNorm",
    "MaxTemperature": "_PCTempMax",
    "MinPressure": "_PCPressMin",
    "NormPressure": "_PCPressNorm",
    "MaxPressure": "_PCPressMax",
    "AreaClassification": "_PCAREA_CLASS_REQ",
    "AmbTemperatureReq": "_PCAmbTempReqs",
    "CageType": "_CageType",
    "CageUpperConnSize": "_BdyCageUpConnSize",
    "CageUpperConnType": "_BdyCageUpConnType",
    "CageLowerConnSize": "_BdyCageLowConnSize",
    "CageLowerConnType": "_BdyCageLowConnType",
    "CageMaterial": "_BdyCageMaterial",
    "CageRating": "_BdyCageRating",
    "CageRotatableHead": "_BdyCageRotatHead",
    "CageOrientation": "_BdyCageOrient",
    "CageDrainValve": "_BdyCageDrainValve",
    "CageVentValve": "_BdyCageVentValve",
    "CageLength": "_CageLen",
    "CageDiameter": "_BdyCageDiam",
    "CageNozzleExtLength": "_BdyCageNozzExtLen",
    "CageNozzleSpan": "_BdyCageSpanbwNozz",
    "CageCoolingExtension": "_BdyCageCoolExt",
    "CageMounting": "CageMounting",
    "TransmitterVoltage": "_XmtrVotage",
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
    "TransmitterAccuracy": "_XmtrAccuracy",
    "TransmitterSensorConn": "_XmtrSensConn",
    "TransmitterGlandConn": "_XmtrGlandConn",
    "TransmitterBodyMaterial": "_trBdyMaterial",
    "TransmitterCalRange": "_XmtrCalibRange",
    "TransmitterFullRange": "_XmtrFullRange",
    "TransmitterGasketMat": "_XmtrGaskMaterial",
    "TransmitterSSTag": "_XmtrSSTag",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "ElementType": "_ElmntType",
    "ElementGauge": "_ElmntGauge",
    "ElementGaugeMat": "_ElmntGaugeMaterial",
    "ElementGaugeDim": "_ElmntGaugeDim",
    "ElementInsertDepth": "_ElmntInsertDepth",
    "ElementProcConn": "_ElmntProcConn",
    "ElementTransmitterConn": "_ElmntXmtrConn",
    "ElementDim": "_ElmntDim",
    "ElementMat": "_ElmntMaterial",
    "ElementRadarFrequency": "_ElmntRadarFreq",
    "ElementDisplacerExt": "_ElmntDisplacExtension",
    "ElementSpringTubeMat": "_ElmntDisplacSprTubMat",
    "ElementFloatWellClearance": "_ElmntFloatWellClear",
    "ElementFloatShaftClearance": "_ElmntFloatShaftClear",
    "SwtchRelay1TagNumber": "_SwtchRel1TagNo",
    "SwtchRelay1PowerRating": "_SwtchRel1PwrRating",
    "SwtchRelay1Type": "_SwtchRel1Type",
    "SwtchRelay2PowerRating": "_SwtchRel2PwrRating",
    "SwtchRelay2Type": "_SwtchRel2Type",
    "SwtchRelay2AlmSetting": "_SwtchRel2AlarmSet",
    "SwtchCtctRating": "_SwtchContRating",
    "SwtchGlandConn": "_SwtchGlandConn",
    "SwtchManufacturer": "_SwtchManufact",
    "SwtchModelNo": "_SwtchModelNo",
    "SwtchElectricalProtection": "_SwtchElectProtect",
    "SwtchTempCat": "_SwtchTempCatgy",
    "SwtchGasGroup": "_SwtchGasGrp",
    "SwtchIP1": "SwtchIP1",
    "SwtchIP2": "_SwtchEnclProtectIP1",
    "Note_Field1": "_Note1",
    "Note_Field2": "_Note2",
    "Note_Field3": "_Note3",
    "Note_Field4": "_Note4",
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

