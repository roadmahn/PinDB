# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_FI(Instrument):
    # PROCESS CONDITIONS
    ProcessFluid = ""
    FluidStatePresent = ""
    Compressibility = ""
    OperatingTemperature = ""
    OperatingPressure = ""
    VaporPressure = ""
    BasePressure = ""
    BaseTemperature = ""
    TemperatureUnits = ""
    PressureUnits = ""
    VaporPressureUnits = ""
    FlangeMaterial = ""
    AreaClassification = ""
    LiquidFlowUnits = ""
    MinLiquidFlow = ""
    NormLiquidFlow = ""
    MaxLiquidFlow = ""
    VapourFlowUnits = ""
    MinVapourFlow = ""
    NormVapourFlow = ""
    MaxVapourFlow = ""
    OperatingLiquidDensity = ""
    LiquidDensityUnits = ""
    OperatingVapourDensity = ""
    VapourDensityUnits = ""
    OperatingViscosity = ""
    ViscosityUnits = ""
    SpecificHeatRatio = ""
    SolidsPercentage = ""
    SteamQuality = ""
    PipeMaterial = ""
    FlangeRating = ""
    AmbTemperatureReq = ""
    # ELEMENT
    ElementType = ""
    ElementShaftMat = ""
    ElementBladesMat = ""
    ElementInstallation = ""
    ElementSealType = ""
    ElementSealMat = ""
    ElementMeterSize = ""
    ElementFaceToFace = ""
    ElementBearingType = ""
    ElementBearingMat = ""
    ElementProcessConn = ""
    ElementGearType = ""
    ElementGearMat = ""
    ElementTxConnection = ""
    ElementPackingType = ""
    ElementTempRating = ""
    ElementPressRating = ""
    ElementCouplingType = ""
    ElementMaxDP = ""
    ElementRatedFlowRange = ""
    ElementAccuracy = ""
    ElementRepeatability = ""
    ElementHousingMat = ""
    ElementRotMat = ""
    ElementStrainerType = ""
    ElementShutOffValve = ""
    ElementFlowLimitValve = ""
    ElementStrainerSize = ""
    ElementTempComp = ""
    ElementTotalizer = ""
    ElementStrainerMat = ""
    ElementAirEliminator = ""
    ElementStrainerMesh = ""
    # FLOW TOTALIZER
    TotalizerRegisterType = ""
    TotalizerDigitNumbers = ""
    TotalizerCapacity = ""
    TotalizerUnits = ""
    TotalizerSetStop = ""
    TotalizerMounting = ""
    TotalizerReset = ""
    # TRANSMITTER
    TransmitterVoltage = ""
    TransmitterSensorConn = ""
    TransmitterPowerWiring = ""
    TransmitterSignalType = ""
    TransmitterGlandConn = ""
    TransmitterProtocol = ""
    TransmitterLocation = ""
    TransmitterIndicate = ""
    TranmitterIsolate = ""
    TransmitterBodyMaterial = ""
    TransmitterElectricalProtection = ""
    TransmitterTempCategory = ""
    TransmitterMounting = ""
    TransmitterSSTag = ""
    TransmitterGasGroup = ""
    TransmitterIPRating = ""
    TransmitterFullRange = ""
    TransmitterFacCalibration = ""
    TransmitterCalRange = ""
    Connection1 = ""
    Connection2 = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
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
    "ProcessFluid": "_PCFluid",
    "FluidStatePresent": "_PCFluidStatePresent",
    "Compressibility": "_PCCompress",
    "OperatingTemperature": "_PCTempAtOp",
    "OperatingPressure": "_PCPressAtOp",
    "VaporPressure": "_PCVapPress",
    "BasePressure": "_PCBasePress",
    "BaseTemperature": "_PCBaseTemp",
    "TemperatureUnits": "_PCTempUnits",
    "PressureUnits": "_PCPressUnits",
    "VaporPressureUnits": "_PCVapPressUnits",
    "FlangeMaterial": "_PCFlngMat",
    "AreaClassification": "_PCAREA_CLASS_REQ",
    "LiquidFlowUnits": "_PCLiqFlowUnits",
    "MinLiquidFlow": "_PCMinLiqFlow",
    "NormLiquidFlow": "_PCNormOpLiqFlow",
    "MaxLiquidFlow": "_PCMaxFullScaleLiqFlow",
    "VapourFlowUnits": "_PCVapFlowUnits",
    "MinVapourFlow": "_PCMinVapFlow",
    "NormVapourFlow": "_PCNormOpVapFlow",
    "MaxVapourFlow": "_PCMaxFullScaleVapFlow",
    "OperatingLiquidDensity": "_PCLiqSgDensityAtOp",
    "LiquidDensityUnits": "_PCLiqSgDensityAtOpUnits",
    "OperatingVapourDensity": "_PCVapSgDensityAtOp",
    "VapourDensityUnits": "_PCVapSgDensityAtOpUnits",
    "OperatingViscosity": "_PCViscosityAtOp",
    "ViscosityUnits": "_PCViscosityAtOpUnits",
    "SpecificHeatRatio": "_PCSpHeatRatio",
    "SolidsPercentage": "_PCPercentSolids",
    "SteamQuality": "_PCSteamPercentQualityOrSuperHeat",
    "PipeMaterial": "_PCPipeMaterial",
    "FlangeRating": "_PCFlngRating",
    "AmbTemperatureReq": "_PCAmbTempReqs",
    "ElementType": "_ElmntType",
    "ElementShaftMat": "_ElmntMatrlShaft",
    "ElementBladesMat": "_ElmntMatrlBlades",
    "ElementInstallation": "_ElmntInstal",
    "ElementSealType": "_ElmntSealType",
    "ElementSealMat": "_ElmntSealMatrl",
    "ElementMeterSize": "_ElmntMeterSize",
    "ElementFaceToFace": "_ElmntFaceFace",
    "ElementBearingType": "_ElmntBearType",
    "ElementBearingMat": "_ElmntBearingMatrl",
    "ElementProcessConn": "_ElmntProcConn",
    "ElementGearType": "_ElmntGearType",
    "ElementGearMat": "_ElmntGearMatrl",
    "ElementTxConnection": "_ElmntTotalTxConn",
    "ElementPackingType": "_ElmntPackType",
    "ElementTempRating": "_ElmntTempRating",
    "ElementPressRating": "_ElmntPressRating",
    "ElementCouplingType": "_ElmntTypeCoup",
    "ElementMaxDP": "_ElmntMaxdP",
    "ElementRatedFlowRange": "_ElmntRatedFlowRange",
    "ElementAccuracy": "_ElmntAccuracy",
    "ElementRepeatability": "_ElmntRepeadability",
    "ElementHousingMat": "_ElmntMatrlHousing",
    "ElementRotMat": "_ElmntMatrlRotElem",
    "ElementStrainerType": "_ElmntStrainerType",
    "ElementShutOffValve": "_ElmntShutOffValve",
    "ElementFlowLimitValve": "_ElmntFlowLimValve",
    "ElementStrainerSize": "_ElmntStrainerSize",
    "ElementTempComp": "_ElmntTempCompensat",
    "ElementTotalizer": "_ElmntTotalizer",
    "ElementStrainerMat": "_ElmntStrainerMatrl",
    "ElementAirEliminator": "_ElmntAirEliminator",
    "ElementStrainerMesh": "_ElmntStrainerMesh",
    "TotalizerRegisterType": "_FTRegisterType",
    "TotalizerDigitNumbers": "_FTNoOfDigits",
    "TotalizerCapacity": "_FTCapacity",
    "TotalizerUnits": "_FTTotalUnits",
    "TotalizerSetStop": "_FTSetStop",
    "TotalizerMounting": "_FTMounting",
    "TotalizerReset": "_FTReset",
    "TransmitterVoltage": "_XmtrVolt",
    "TransmitterSensorConn": "_XmtrSensConn",
    "TransmitterPowerWiring": "_XmtrPwrWiring",
    "TransmitterSignalType": "_XmtrSigType",
    "TransmitterGlandConn": "_XmtrGlandConn",
    "TransmitterProtocol": "_XmtrCommProtocol",
    "TransmitterLocation": "_XmtrLocation",
    "TransmitterIndicate": "_XmtrIndicate",
    "TranmitterIsolate": "_XmtrIsolate",
    "TransmitterBodyMaterial": "_XmtrBdyMaterial",
    "TransmitterElectricalProtection": "_PROT_TYPE",
    "TransmitterTempCategory": "_TEMP_CLASS",
    "TransmitterMounting": "_XmtrMount",
    "TransmitterSSTag": "_XmtrSSTag",
    "TransmitterGasGroup": "_GASGROUP",
    "TransmitterIPRating": "_IP_RATING",
    "TransmitterFullRange": "_XmtrFullRange",
    "TransmitterFacCalibration": "_XmtrFactCalib",
    "TransmitterCalRange": "_XmtrCalibRange",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Remarks": "_Remarks",
    "EC_TE_Certificate": "_EC_TYP_EX_CERT",
    "EC_DoC": "_EC_DECL_CONF",
    "QA_Notification": "_PROD_QA_NOTIFICATION",
    "QA_Not_Date": "_PROD_QA_NOTIFICATION",
    "NOBO_Number": "_NOBO_REF",
    "User3": "_User3",
    "User4": "_User4",
    "User5": "_User5",
    "User6": "_User6",
    "User7": "_User7",
    "User8": "_User8",
    "User9": "_User9",
    }


