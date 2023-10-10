# -*- coding: utf-8 -*-
from Classes.ElecEquipment import ElecEquipment

class DS_JIP33_AC_UPS(ElecEquipment):
    # GENERAL
    
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
    # TRANSMITTER
    XmtrVolt = ""
    XmtrPwrWiring = ""
    XmtrSigType = ""
    XmtrCommProtocol = ""
    XmtrLocation = ""
    XmtrSmart = ""
    XmtrIndicate = ""
    XmtrIsolate = ""
    XmtrElecProt = ""
    XmtrTempCat = ""
    XmtrGasGroup = ""
    XmtrIPRating = ""
    XmtrAmbTempCompens = ""
    XmtrFactCalib = ""
    XmtrTCBurnout = ""
    XmtrRTDConst = ""
    XmtrRange = ""
    XmtrBdyMat = ""
    XmtrSSTag = ""
    XmtrAccuracy = ""
    # ELEMENT
    ElmntTagName = ""
    ElmntType = ""
    ElmntProbeType = ""
    ElmntSingleDuplex = ""
    ElmntFixAdjust = ""
    ElmntLims = ""
    ElmntSensLen = ""
    ElmntSheathDiam = ""
    ElmntSheathMat = ""
    ElmntNoOfLeadWireTerm = ""
    ElmntLeadLen = ""
    ElmntSpringLoad = ""
    ElmntGrndJunction = ""
    ElmntProcConn = ""
    ElmntSigConn = ""
    ElmntCodeStands = ""
    ElmntManufact = ""
    ElmntModel = ""
    # CONNECTION HEAD
    CHElectProtect = ""
    CHTempCatgy = ""
    CHGasGrp = ""
    CHEnclProtectIP1 = ""
    CHEnclProtectIP2 = ""
    CHMat = ""
    CHStyle = ""
    CHSensConn = ""
    CHGlandConn = ""
    CHTermStrip = ""
    CHManufact = ""
    CHModel = ""
    # CONNECTIONS
    Connection1 = ""
    Connection2 = ""
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
    "MinPress": "_PCMinPress",
    "NormPress": "_PCNormPress",
    "MaxPress": "_PCMaxPress",
    "PressUnits": "_PressUnits",
    "MinTemp": "_PCMinTemp",
    "NormTemp": "_PCNormTemp",
    "MaxTemp": "_PCMaxTemp",
    "TempUnits": "_TempUnits",
    "MinFlow": "_PCMinFlow",
    "NormFlow": "_PCNormFlow",
    "MaxFlow": "_PCMaxFlow",
    "FlowUnits": "_FlowUnits",
    "Fluid": "_PCFluid",
    "AreaClassification": "_AREA_CLASS_REQ",
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
    "XmtrVolt": "_XmtrVolt",
    "XmtrPwrWiring": "_XmtrPwrWiring",
    "XmtrSigType": "_XmtrSigType",
    "XmtrCommProtocol": "_XmtrCommProtocol",
    "XmtrLocation": "_XmtrLocation",
    "XmtrSmart": "_XmtrSmart",
    "XmtrIndicate": "_XmtrIndicate",
    "XmtrIsolate": "_XmtrIsolate",
    "XmtrElecProt": "_PROT_TYPE",
    "XmtrTempCat": "_TEMP_CLASS",
    "XmtrGasGroup": "_GASGROUP",
    "XmtrIPRating": "_IP_RATING",
    "XmtrAmbTempCompens": "_XmtrAmbTempCompens",
    "XmtrFactCalib": "_XmtrFactCalib",
    "XmtrTCBurnout": "_XmtrTCBurnout",
    "XmtrRTDConst": "_XmtrRTDConst",
    "XmtrRange": "_Range",
    "XmtrBdyMat": "_XmtrBdyMat",
    "XmtrSSTag": "_XmtrSSTag",
    "XmtrAccuracy": "_XmtrAccuracy",
    "ElmntTagName": "_ElmntTagName",
    "ElmntType": "_ElmntType",
    "ElmntProbeType": "_ElmntProbeType",
    "ElmntSingleDuplex": "_ElmntSingleDuplex",
    "ElmntFixAdjust": "_ElmntFixAdjust",
    "ElmntLims": "_ElmntLims",
    "ElmntSensLen": "_ElmntSensLen",
    "ElmntSheathDiam": "_ElmntSheathDiam",
    "ElmntSheathMat": "_ElmntSheathMat",
    "ElmntNoOfLeadWireTerm": "_ElmntNoOfLeadWireTerm",
    "ElmntLeadLen": "_ElmntLeadLen",
    "ElmntSpringLoad": "_ElmntSpringLoad",
    "ElmntGrndJunction": "_ElmntGrndJunction",
    "ElmntProcConn": "_ElmntProcConn",
    "ElmntSigConn": "_ElmntSigConn",
    "ElmntCodeStands": "_ElmntCodeStands",
    "ElmntManufact": "_ElmntManufact",
    "ElmntModel": "_ElmntModel",
    "CHElectProtect": "_CHElectProtect",
    "CHTempCatgy": "_CHTempCatgy",
    "CHGasGrp": "_CHGasGrp",
    "CHEnclProtectIP1": "_CHEnclProtectIP1",
    "CHEnclProtectIP2": "_CHEnclProtectIP2",
    "CHMat": "_CHMat",
    "CHStyle": "_CHStyle",
    "CHSensConn": "_CHSensConn",
    "CHGlandConn": "_CHGlandConn",
    "CHTermStrip": "_CHTermStrip",
    "CHManufact": "_CHManufact",
    "CHModel": "_CHModel",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
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


