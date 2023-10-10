# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_PIG(Instrument):
    DetectionType = ""
    Location = ""
    AreaClassification = ""
    AmbRelativeHumidity = ""
    AmbTemperature = ""
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
    TransmitterTempComp = ""
    TransmitterFacCalibration = ""
    TransmitterSensorConn = ""
    TransmitterGlandConn = ""
    TransmitterApprovals = ""
    TransmitterBodyMaterial = ""
    TransmitterMounting = ""
    TransmitterContactRating = ""
    TransmitterFullRange = ""
    TransmitterCalRange = ""
    TransmitterSSTag = ""
    TransmitterRelay1TagNumber = ""
    TransmitterRelay1Voltage = ""
    TransmitterRelay1Type = ""
    TransmitterRelay1AlmSetting = ""
    TransmitterRelay2TagNumber = ""
    TransmitterRelay2Voltage = ""
    TransmitterRelay2Type = ""
    TransmitterRelay2AlmSetting = ""
    TransmitterRelay3TagNumber = ""
    TransmitterRelay3Voltage = ""
    TransmitterRelay3Type = ""
    TransmitterRelay3AlmSetting = ""
    TransmitterRelay4TagNumber = ""
    TransmitterRelay4Voltage = ""
    TransmitterRelay4Type = ""
    TransmitterRelay4AlmSetting = ""
    Connection1 = ""
    Connection2 = ""
    # ELEMENT
    ElementMounting = ""
    ElementAccuracy = ""
    ElementBodyMat = ""
    ElementMaxTemp = ""
    ElementProcConn = ""
    ElementRatedFlow = ""
    ElementPipingMat = ""
    ElementManufacturer = ""
    ElemenMaxPress = ""
    ElementModelNo = ""
    ElementRepeatability = ""
    # ACCESSORY
    JunctionBox = ""
    RainShield = ""
    CalibrationEquipment = ""
    RemoteMountingKit = ""
    DustCover = ""
    DuctMountAssembly = ""
    Other1 = ""
    Other2 = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Note_Field4 = ""
    Note_Field5 = ""
    Note_Field6 = ""
    Note_Field7 = ""
    Remarks = ""
    # ATEX  REGISTER
    EC_TE_Certificate = ""
    EC_DoC = ""
    QA_Notification = ""
    QA_Not_Date = ""
    NOBO_Number = ""
    User3 = ""
    User4 = ""
    InternalFieldsDict = {
    "DetectionType": "_GenDetType",
    "Location": "_GenLocat",
    "AreaClassification": "_AREA_CLASS_REQ",
    "AmbRelativeHumidity": "_GenAmbRelHumid",
    "AmbTemperature": "_GenAmbTemp",
    "TransmitterVoltage": "_XmtrVolt",
    "TransmitterPowerWiring": "_XmtrPowWiring",
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
    "TransmitterTempComp": "_XmtrAmbTempCompenst",
    "TransmitterFacCalibration": "_XmtrFactCalib",
    "TransmitterSensorConn": "_XmtrSensConn",
    "TransmitterGlandConn": "_XmtrGlandConn",
    "TransmitterApprovals": "_XmtrApproval",
    "TransmitterBodyMaterial": "_XmtrBdyMat",
    "TransmitterMounting": "_XmtrMount",
    "TransmitterContactRating": "_XmtrContrating",
    "TransmitterFullRange": "_XmtrFullRange",
    "TransmitterCalRange": "_XmtrCalibRange",
    "TransmitterSSTag": "_XmtrSSTag",
    "TransmitterRelay1TagNumber": "_XmtrRelay1TagNo",
    "TransmitterRelay1Voltage": "_XmtrRelay1Volt",
    "TransmitterRelay1Type": "_XmtrRelay1Type",
    "TransmitterRelay1AlmSetting": "_XmtrRelay1AlarmSetting",
    "TransmitterRelay2TagNumber": "_XmtrRelay2TagNo",
    "TransmitterRelay2Voltage": "_XmtrRelay2Volt",
    "TransmitterRelay2Type": "_XmtrRelay2Type",
    "TransmitterRelay2AlmSetting": "_XmtrRelay2AlarmSetting",
    "TransmitterRelay3TagNumber": "_XmtrRelay3TagNo",
    "TransmitterRelay3Voltage": "_XmtrRelay3Volt",
    "TransmitterRelay3Type": "_XmtrRelay3Type",
    "TransmitterRelay3AlmSetting": "_XmtrRelay3AlarmSetting",
    "TransmitterRelay4TagNumber": "_XmtrRelay4TagNo",
    "TransmitterRelay4Voltage": "_XmtrRelay4Volt",
    "TransmitterRelay4Type": "_XmtrRelay4Type",
    "TransmitterRelay4AlmSetting": "_XmtrRelay4AlarmSetting",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "ElementMounting": "_ElmntMountType",
    "ElementAccuracy": "_ElmntAccuracy",
    "ElementBodyMat": "_ElmntBdyMat",
    "ElementMaxTemp": "_ElmntMaxTemp",
    "ElementProcConn": "_ElmntProcConn",
    "ElementRatedFlow": "_ElmntRatedFlowRange",
    "ElementPipingMat": "_ElmntProcPipingMat",
    "ElementManufacturer": "_ElmntManufact",
    "ElemenMaxPress": "_ElmntMaxPress",
    "ElementModelNo": "_ElmntModelNo",
    "ElementRepeatability": "_ElmntRepeatability",
    "JunctionBox": "_AccJuncBox",
    "RainShield": "_AccRainShld",
    "CalibrationEquipment": "_AccCalibEquip",
    "RemoteMountingKit": "_AccRmtMntKit",
    "DustCover": "_AccDustCover",
    "DuctMountAssembly": "_AccDuctMntAssmbl",
    "Other1": "_AccOther1",
    "Other2": "_AccOther2",
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Note_Field4": "_Notes4",
    "Note_Field5": "_Notes5",
    "Note_Field6": "_Notes6",
    "Note_Field7": "_Notes7",
    "Remarks": "_Remarks",
    "EC_TE_Certificate": "_EC_TYP_EX_CERT",
    "EC_DoC": "_EC_DECL_CONF",
    "QA_Notification": "_PROD_QA_NOTIFICATION",
    "QA_Not_Date": "_PROD_QA_NOT_DATE",
    "NOBO_Number": "_NOBO_REF",
    "User3": "_User3",
    "User4": "_User4"
    }

    

