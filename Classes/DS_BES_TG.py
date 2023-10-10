# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_TG(Instrument):
    # PROCESS CONDITIONS
    MinPress = ""
    NormPress = ""
    MaxPress = ""
    PressUnits = ""
    MinTemp = ""
    NormTemp = ""
    MaxTemp = ""
    TempUnits = ""
    MinFlow = ""
    NormFlow = ""
    MaxFlow = ""
    FlowUnits = ""
    Fluid = ""
    AreaClassification = ""
    PCAmbTempReqs = ""
    Service = ""
    Critic = ""
    # THERMOWELL
    ThermTagNo = ""
    ThermType = ""
    ThermNACE = ""
    ThermProcConn = ""
    ThermElmntConn = ""
    ThermVibrCalc = ""
    ThermMaxPressRating = ""
    ThermMaterial = ""
    ThermWellID = ""
    ThermWellOD = ""
    ThermInsertLen = ""
    ThermTipDiam = ""
    ThermStemLen = ""
    ThermTipThick = ""
    ThermRootDiam = ""
    ThermHeadLen = ""
    ThermLAGLen = ""
    ThermImmersLen = ""
    ThermManufact = ""
    ThermModel = ""
    # BI METAL ELEMENT
    BiMetElmntMat = ""
    BiMetalElmntStemLen = ""
    BiMetalElmntStemDia = ""
    BiMetalElmntStemMat = ""
    BiMetalTempRange = ""
    BiMetalProcWellConn = ""
    BiMetalGaugeConn = ""
    BiMetalCodeStandard = ""
    BiMetalManufact = ""
    BiMetalModel = ""
    # FLUID FILLED ELEMENT
    FFEFillFluid = ""
    FFECodeStnrd = ""
    FFEBulbType = ""
    FFEBulbMat = ""
    FFEBulbDiam = ""
    FFEBulbLen = ""
    FFECapillaryMat = ""
    FFECapillaryLen = ""
    FFECapillaryArmour = ""
    FFEProcWellConn = ""
    FFEGaugeConn = ""
    FFESAMAClass = ""
    FFEManufact = ""
    FFEModel = ""
    # GAUGE
    GaugeType = ""
    GaugeExtReset = ""
    GaugeMount = ""
    GaugeHermetSealed = ""
    GaugeAntiParallax = ""
    GaugeSize = ""
    GaugeCaseMat = ""
    GaugeGlassMat = ""
    GaugeRange = ""
    GaugeScale = ""
    GaugeBezel = ""
    GaugeAccuracy = ""
    GaugeOverRange = ""
    GaugeDampening = ""
    GaugeZerAdjust = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Note_Field5 = ""
    Note_Field6 = ""
    Note_Field7 = ""
    Note_Field8 = ""
    Note_Field9 = ""
    Note_Field10 = ""
    Note_Field11 = ""
    Remarks = ""
    InternalFieldsDict = {
    "MinPress": "_PCMinPress",
    "NormPress": "_PCNormPress",
    "MaxPress": "_PCMaxPress",
    "PressUnits": "_PCPressUnits",
    "MinTemp": "_PCMinTemp",
    "NormTemp": "_PCNormTemp",
    "MaxTemp": "_PCMaxTemp",
    "TempUnits": "_PCTempUnits",
    "MinFlow": "_PCMinFlow",
    "NormFlow": "_PCNormFlow",
    "MaxFlow": "_PCMaxFlow",
    "FlowUnits": "_PCFlowUnits",
    "Fluid": "_PCFluid",
    "AreaClassification:": "_PCAREA_CLASS_REQ",
    "PCAmbTempReqs": "_PCAmbTempReqs",
    "Service": "_PCService",
    "Critic": "_PCCritic",
    "ThermTagNo": "_ThermTagNo",
    "ThermType": "_ThermType",
    "ThermNACE": "_ThermNACE",
    "ThermProcConn": "_ThermProcConn",
    "ThermElmntConn": "_ThermElmntConn",
    "ThermVibrCalc": "_ThermVibrCalc",
    "ThermMaxPressRating": "_ThermMaxPressRating",
    "ThermMaterial": "_ThermMaterial",
    "ThermWellID": "_ThermWellID",
    "ThermWellOD": "_ThermWellOD",
    "ThermInsertLen": "_ThermInsertLen",
    "ThermTipDiam": "_ThermTipDiam",
    "ThermStemLen": "_ThermStemLen",
    "ThermTipThick": "_ThermTipThick",
    "ThermRootDiam": "_ThermRootDiam",
    "ThermHeadLen": "_ThermHeadLen",
    "ThermLAGLen": "_ThermLAGLen",
    "ThermImmersLen": "_ThermImmersLen",
    "ThermManufact": "_ThermManufact",
    "ThermModel": "_ThermModel",
    "BiMetElmntMat": "_BiMetElmntMat",
    "BiMetalElmntStemLen": "_BiMetalElmntStemLen",
    "BiMetalElmntStemDia": "_BiMetalElmntStemDia",
    "BiMetalElmntStemMat": "_BiMetalElmntStemMat",
    "BiMetalTempRange": "_BiMetalTempRange",
    "BiMetalProcWellConn": "_BiMetalProcWellConn",
    "BiMetalGaugeConn": "_BiMetalGaugeConn",
    "BiMetalCodeStandard": "_BiMetalCodeStandard",
    "BiMetalManufact": "_BiMetalManufact",
    "BiMetalModel": "_BiMetalModel",
    "FFEFillFluid": "_FFEFillFluid",
    "FFECodeStnrd": "_FFECodeStnrd",
    "FFEBulbType": "_FFEBulbType",
    "FFEBulbMat": "_FFEBulbMat",
    "FFEBulbDiam": "_FFEBulbDiam",
    "FFEBulbLen": "_FFEBulbLen",
    "FFECapillaryMat": "_FFECapillaryMat",
    "FFECapillaryLen": "_FFECapillaryLen",
    "FFECapillaryArmour": "_FFECapillaryArmour",
    "FFEProcWellConn": "_FFEProcWellConn",
    "FFEGaugeConn": "_FFEGaugeConn",
    "FFESAMAClass": "_FFESAMAClass",
    "FFEManufact": "_FFEManufact",
    "FFEModel": "_FFEModel",
    "GaugeType": "_GaugeType",
    "GaugeExtReset": "_GaugeExtReset",
    "GaugeMount": "_GaugeMount",
    "GaugeHermetSealed": "_GaugeHermetSealed",
    "GaugeAntiParallax": "_GaugeAntiParallax",
    "GaugeSize": "_GaugeSize",
    "GaugeCaseMat": "_GaugeCaseMat",
    "GaugeGlassMat": "_GaugeGlassMat",
    "GaugeRange": "_Range",
    "GaugeScale": "_GaugeScale",
    "GaugeBezel": "_GaugeBezel",
    "GaugeAccuracy": "_GaugeAccuracy",
    "GaugeOverRange": "_GaugeOverRange",
    "GaugeDampening": "_GaugeDampening",
    "GaugeZerAdjust": "_GaugeZerAdjust",
    "Note_Field1": "_Note1",
    "Note_Field2": "_Note2",
    "Note_Field3": "_Note3",
    "Note_Field4": "_Note4",
    "Note_Field5": "_Note5",
    "Note_Field6": "_Note6",
    "Note_Field7": "_Note7",
    "Note_Field8": "_Note8",
    "Note_Field9": "_Note9",
    "Note_Field10": "_Note10",
    "Note_Field11": "_Note11",
    "Remarks": "_Remarks"
    }

    
