# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_LG(Instrument):
    # PROCESS CONDITIONS
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
    # BODY / CAGE
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
    # ELEMENT
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
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
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
    "Note_Field1": "_Note1",
    "Note_Field2": "_Note2",
    "Note_Field3": "_Note3",
    "Note_Field4": "_Note4",
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


