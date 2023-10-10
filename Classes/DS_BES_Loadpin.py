# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Loadpin(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    Function = ""
    # SPECIFICATIONS    
    RatedLoad = ""
    ProofLoad = ""
    Accuracy = ""
    SafetyFactor = ""
    OperTemperature = ""
    StorageTemperature = ""
    OutputSignal = ""
    ElectricalConnection = ""
    CableConnection = ""
    CableLength = ""
    ExcitationVoltage = ""
    MaxExcitationVoltage = ""
    SupplyVoltage = ""
    MaxSupplyVoltage = ""
    BridgeResistance = ""
    Sensitivity = ""
    DataLogging = ""
    BatteryType = ""
    ActiveBatteryLife = ""
    StandbyBatteryLife = ""
    LoggingBatteryLife = ""
    TelemetryFrequency = ""
    SystemRange = ""
    SamplingRange = ""
    Length = ""
    Diameter = ""
    Connection1 = ""
    Connection2 = ""
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
    "RatedLoad": "_SpecRatedLoad",
    "ProofLoad": "_SpecProofLoad",
    "Accuracy": "_SpecAccuracy",
    "SafetyFactor": "_SpecSafetyFactor",
    "OperTemperature": "_SpecOperTemperature",
    "StorageTemperature": "_SpecStoreTemperature",
    "OutputSignal": "_SpecOutputSignal",
    "ElectricalConnection": "_SpecElectrConnection",
    "CableConnection": "_SpecCableConnection",
    "CableLength": "_SpecCableLength",
    "ExcitationVoltage": "_SpecExcitationVoltage",
    "MaxExcitationVoltage": "_SpecMaxExcitVoltage",
    "SupplyVoltage": "_SpecSupplyVoltage",
    "MaxSupplyVoltage": "_SpecMaxSupplyVoltage",
    "BridgeResistance": "_SpecBridgeResistance",
    "Sensitivity": "_SpecSensitivity",
    "DataLogging": "_SpecDataLogging",
    "BatteryType": "_SpecBatteryType",
    "ActiveBatteryLife": "_SpecActiveBatteryLife",
    "StandbyBatteryLife": "_SpecStdbyBatteryLife",
    "LoggingBatteryLife": "_SpecLoggingBatteryLife",
    "TelemetryFrequency": "_SpecTelemFrequency",
    "SystemRange": "_SpecSystemRange",
    "SamplingRange": "_SpecSamplingRange",
    "Length": "_SpecLength",
    "Diameter": "_SpecDiameter",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",
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

    