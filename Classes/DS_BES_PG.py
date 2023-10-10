# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_PG(Instrument):
    # PROCESS CONDITIONS
    HiPressureFluid = ""
    HiMinPressure = ""
    HiNormPressure = ""
    HiMaxPressure = ""
    HiMinTemperature = ""
    HiNormTemperature = ""
    HiMaxTemperature = ""
    HiService = ""
    HISolids = ""
    HiQuality = ""
    LoPressureFluid = ""
    LoMinPressure = ""
    LoNormPressure = ""
    LoMaxPressure = ""
    LoMinTemperature = ""
    LoNormTemperature = ""
    LoMaxTemperature = ""
    LoService = ""
    LoSolids = ""
    LoQuality = ""
    AreaClassification = ""
    AmbTemperatureReq = ""
    PressUnits = ""
    TempUnits = ""
    # ELEMENT
    ElementType = ""
    ElementDiaphragmMat = ""
    ElementPressMinSpan = ""
    ElementPressMaxSpan = ""
    ElementMaxPressure = ""
    ElementFluidMat = ""
    # GAUGE
    GaugeMounting = ""
    GaugeLiquidFilled = ""
    GaugeDialDiam = ""
    GaugeDialColor = ""
    GaugeCaseMat = ""
    GaugeLensMat = ""
    GaugeRingType = ""
    GaugeUnits = ""
    GaugeRange = ""
    GaugeNominalAccuracy = ""
    GaugeMovementMat = ""
    GaugeSnubberOpt = ""
    GaugeSyphonOpt = ""
    GaugeSocketMat = ""
    GaugeBOP = ""
    GaugePLV = ""
    GaugeLoPressureConn = ""
    GaugeHiPressureConn = ""
    GaugeAccessoires = ""
     # CAPILLARY
    CapillaryLength = ""
    CapillaryID = ""
    CapillaryArmor = ""
    CapillaryFillFluid = ""
    CapillaryResponseTime = ""
    CapillarySGat60oF = ""
    CapillaryMat = ""
    # SEAL DIAPHRAM & CAPILLARY
    SealDiaHiPressSizeType = ""
    SealDiaHiPressThickness = ""
    SealDialHiPressMat = ""
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
    MnfldModel = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Note_Field5 = ""
    Note_Field6 = ""
    Remarks = ""
    User3 = ""
    User4 = ""
    User5 = ""
    User6 = ""
    User7 = ""
    User8 = ""
    User9 = ""
    User10 = ""
    InternalFieldsDict = {
    "HiPressureFluid": "_PCHiPressFluid",
    "HiMinPressure": "_PCHiMinPress",
    "HiNormPressure": "_PCHiNormPress",
    "HiMaxPressure": "_PCHiMaxPress",
    "HiMinTemperature": "_PCHiPressMinTemp",
    "HiNormTemperature": "_PCHiPressNormTemp",
    "HiMaxTemperature": "_PCHiPressMaxTemp",
    "HiService": "_PCHiPressService",
    "HISolids": "_PCHiPressPercSolids",
    "HiQuality": "_PCHiPressPercQuality",
    "LoPressureFluid": "_PCLowPressFluid",
    "LoMinPressure": "_PCLowMinPress",
    "LoNormPressure": "_PCLowNormPress",
    "LoMaxPressure": "_PCLowMaxPress",
    "LoMinTemperature": "_PCLowPressMinTemp",
    "LoNormTemperature": "_PCLowPressNormTemp",
    "LoMaxTemperature": "_PCLowPressMaxTemp",
    "LoService": "_PCLowPressService",
    "LoSolids": "_PCLowPressPercSolids",
    "LoQuality": "_PCLowPressPercQuality",
    "AreaClassification": "_PCAREA_CLASS_REQ",
    "AmbTemperatureReq": "_PCAmbTempReqs",
    "PressUnits": "_PCPressUnits",
    "TempUnits": "_PCTempUnits",
    "ElementType": "_ElmntType",
    "ElementDiaphragmMat": "_ElmntDiaphWettMat",
    "ElementPressMinSpan": "_ElmntPressMinSpan",
    "ElementPressMaxSpan": "_ElmntPressMaxSpan",
    "ElementMaxPressure": "_ElmntMaxPressLimit",
    "ElementFluidMat": "_ElmntFluidMat",
    "GaugeMounting": "_GaugeMount",
    "GaugeLiquidFilled": "_GaugLiqFilled",
    "GaugeDialDiam": "_GaugDialDiam",
    "GaugeDialColor": "_GaugDialColor",
    "GaugeCaseMat": "_GaugeCaseMat",
    "GaugeLensMat": "_GaugeLensMat",
    "GaugeRingType": "_GaugRingType",
    "GaugeUnits": "_GaugUnits",
    "GaugeRange": "_RANGE",
    "GaugeNominalAccuracy": "_GaugNomAccuracy",
    "GaugeMovementMat": "_GaugeMovMaterial",
    "GaugeSnubberOpt": "_GaugSnubOption",
    "GaugeSyphonOpt": "_GaugeSyphOption",
    "GaugeSocketMat": "_GaugeSocketMat",
    "GaugeBOP": "_GaugBlowOutProtect",
    "GaugePLV": "_GaugPressLimValv",
    "GaugeLoPressureConn": "_GaugeProcConnLowPress",
    "GaugeHiPressureConn": "_GaugeProcConnHiPress",
    "GaugeAccessoires": "_GaugeAccessories",
    "CapillaryLength": "_CapillaryLen",
    "CapillaryID": "_CapillaryID",
    "CapillaryArmor": "_CapillaryArmor",
    "CapillaryFillFluid": "_CapillaryFillFluid",
    "CapillaryResponseTime": "_CapillaryMaxRespTime",
    "CapillarySGat60oF": "_CapillarySGAt60F",
    "CapillaryMat": "_CapillaryMat",
    "SealDiaHiPressSizeType": "_SDCHiPressSizeAndType",
    "SealDiaHiPressThickness": "_SDCHiPressThickness",
    "SealDialHiPressMat": "_SDCHiPressMaterial",
    "SealDialHiPressFlushRing": "_SDCHiPressFlushRing",
    "SealDiaLoPressSizeType": "_SDCLowPressSizeAndType",
    "SealDiaLoPressThickness": "_SDCLowPressThickness",
    "SealDialLoPressMat": "_SDCLowPressMaterial",
    "SealDialLoPressFlushRing": "_SDCLowPressFlushRing",
    "SealDiaTempRating": "_SDCTempRating",
    "SealDiaMaxTemp": "_SDCMaxTemp",
    "SealDIaPressRating": "_SDCPressRat",
    "SealDiaMaxPress": "_SDCMaxPress",
    "SealDIaManufacturer": "_SDCManufact",
    "SealDiaModel": "_SDCModel",
    "MnfldType": "_ManifoldType",
    "MnfldMat": "_ManifoldMat",
    "MnfldTransmitterConn": "_ManifoldXmtrConn",
    "MnfldModel": "_ManifoldModel",
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Note_Field4": "_Notes4",
    "Note_Field5": "_Notes5",
    "Note_Field6": "_Notes6",
    "Remarks": "_Remarks",
    "User3": "_User3",
    "User4": "_User4",
    "User5": "_User5",
    "User6": "_User6",
    "User7": "_User7",
    "User8": "_User8",
    "User9": "_User9",
    "User10": "_User10"
    }


