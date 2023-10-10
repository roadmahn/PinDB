# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_GD(Instrument):
    # MONITOR CONDITIONS
    DetectorType = ""
    MCLocation = ""
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    Gas1 = ""
    Gas2 = ""
    Gas3 = ""
    Gas4 = ""
    MolecularWeight1 = ""
    MolecularWeight2 = ""
    MolecularWeight3 = ""
    MolecularWeight4 = ""
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
    TransmitterTempCompensation = ""
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
    ElementTagNumber = ""
    ElementType = ""
    ElementZeroDrift = ""
    ElementSpanDrift = ""
    ElementRangeLimits = ""
    ElementLifespan = ""
    ElementModel = ""
    ElementSensitivity = ""
    ElementResponseTime = ""
    ElementRecoveryTime = ""
    ElementZeroCalibCheck = ""
    ElementSpanCalibCheck = ""
    ElementAccuracy = ""
    CSChemical1 = ""
    CSChemical2 = ""
    CSChemical3 = ""
    CSChemical4 = ""
    InducedError1 = ""
    InducedError2 = ""
    InducedError3 = ""
    InducedError4 = ""
    # ACCESSORY
    JunctionBox = ""
    CalibrationEquipment = ""
    DustCover = ""
    DuctMountAssembly = ""
    RainShield = ""
    RemoteMountingKit = ""
    Other1 = ""
    Other2 = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
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
    "DetectorType": "_MCDetType",
    "MCLocation": "_MCLocat",
    "AreaClassification": "_AREA_CLASS_REQ",
    "AmbientRelHumidity": "_MCAmbRelHumidity",
    "AmbientTemperature": "_MCAmbTempReq",
    "Gas1": "_Gas1",
    "Gas2": "_Gas2",
    "Gas3": "_Gas3",
    "Gas4": "_Gas4",
    "MolecularWeight1": "_Mwt1",
    "MolecularWeight2": "_Mwt2",
    "MolecularWeight3": "_Mwt3",
    "MolecularWeight4": "_Mwt4",
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
    "TransmitterTempCompensation": "_XmtrAmbTempCompenst",
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
    "Connection1":"_Connection1",
    "Connection2":"_Connection2",    
    "ElementTagNumber": "_ElmntTagNo",
    "ElementType": "_ElmntType",
    "ElementZeroDrift": "_ElmntZeroDrft",
    "ElementSpanDrift": "_ElmntSpanDrft",
    "ElementRangeLimits": "_ElmntRnglmts",
    "ElementLifespan": "_ElmntLifeSpan",
    "ElementModel": "_ElmntModel",
    "ElementSensitivity": "_ElmntSenstv",
    "ElementResponseTime": "_ElmntRespTime",
    "ElementRecoveryTime": "_ElmntRecTime",
    "ElementZeroCalibCheck": "_ElmntZeroCalibChk",
    "ElementSpanCalibCheck": "_ElmntSpanCalibChk",
    "ElementAccuracy": "_ElmntAccAtCalibRng",
    "CSChemical1": "_CSChem1",
    "CSChemical2": "_CSChem2",
    "CSChemical3": "_CSChem3",
    "CSChemical4": "_CSChem4",
    "InducedError1": "_IndErr1",
    "InducedError2": "_IndErr2",
    "InducedError3": "_IndErr3",
    "InducedError4": "_IndErr4",
    "JunctionBox": "_AccJuncBox",
    "CalibrationEquipment": "_AccCalibEquip",
    "DustCover": "_AccDustCover",
    "DuctMountAssembly": "_AccDuctMntAssmbl",
    "RainShield": "_AccRainShld",
    "RemoteMountingKit": "_AccRmtMntKit",
    "Other1": "_AccOther1",
    "Other2": "_AccOther2",
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
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

