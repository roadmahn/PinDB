# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Oceanograph(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    # OCEANOGRAPH = ""
    # GENERIC = ""
    SamplingRange = ""
    MinimumChannelWidth = ""
    MulticellVelocityProfiling = ""
    SmartPulseHD = ""
    CompassTilt = ""
    InternalNonvolatileMemory = ""
    # ACOUSTICS = ""
    HorizontalBeamWidth = ""
    VerticalBeamWidth = ""
    SideLobeSuppression = ""
    # WATER VELOCITY = ""
    WVRange = ""
    WVResolution = ""
    WVAccuracy = ""
    # WATER LEVEL = ""
    WLVerticalBeamRange = ""
    WLVerticalBeamAccuracy = ""
    WLPressureSensorRange = ""
    WLPressureSensorAccuracy = ""
    WLWaveHeightSpectra = ""
    # POWER = ""
    InputVoltage = ""
    PowerConsumption = ""
    InputCurrent = ""
    # PHYSICAL PROPERTIES = ""
    WeightinAir = ""
    WeightinWater = ""
    MaxDepth = ""
    MountingPlateDimensions = ""
    OperatingTemperature = ""
    StorageTemperature = ""
    # COMMUNICATION = ""
    StandardProtocols = ""
    Software = ""
    ModbusInterfaceModule = ""
    AnalogOutputOption = ""
    FlowDisplayType = ""
    # TEMP SENSOR = ""
    TempSensorRange = ""
    TempSensorResolution = ""
    TempSensorAccuracy = ""
    Connection1 = ""
    Connection2 = ""    
    ElectricalProtection = ""
    GasGroup = ""
    TempCategory = ""
    IPRating = ""
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
    "ActivationSource": "_GenActivSrc",
    "Function": "_GenFunct",
    "SamplingRange": "_OCSamplingRange",
    "MinimumChannelWidth": "_OCMinChannelWidth",
    "MulticellVelocityProfiling": "_OCMultiCellVelocityProfile",
    "SmartPulseHD": "_OCSmartPulseHD",
    "CompassTilt": "_OCCompassOrTilt",
    "InternalNonvolatileMemory": "_OCInternalMemory",
    "HorizontalBeamWidth": "_OCHorizonBeamWidth",
    "VerticalBeamWidth": "_OCVerticalBeamWidth",
    "SideLobeSuppression": "_OCSideLobeSuppression",
    "WVRange": "_OCWaterVelocityRange",
    "WVResolution": "_OCWaterVelocityResolution",
    "WVAccuracy": "_OCWaterVelocityAccuracy",
    "WLVerticalBeamRange": "_OCWLVerticalBeamAccuracy",
    "WLVerticalBeamAccuracy": "_OCWLVerticalBeamRange",
    "WLPressureSensorRange": "_OCWLPressureSensorRange",
    "WLPressureSensorAccuracy": "_OCWLPressureSensorAccuracy",
    "WLWaveHeightSpectra": "_OCWLWaveHeightSpectra",
    "InputVoltage": "_OCInputVoltage",
    "PowerConsumption": "_OCPowerConsumption",
    "InputCurrent": "_OCInputCurrent",
    "WeightinAir": "_OCWeightInAir",
    "WeightinWater": "_OCWeightInWater",
    "MaxDepth": "_OCMaxDepth",
    "MountingPlateDimensions": "_OCMountingPlateDims",
    "OperatingTemperature": "_OCOperatingTemperature",
    "StorageTemperature": "_OCStorageTemperature",
    "StandardProtocols": "_OCStandardProtocols",
    "Software": "_OCSoftware",
    "ModbusInterfaceModule": "_OCModbusInterfaceModule",
    "AnalogOutputOption": "_OCAnalogOutput",
    "FlowDisplayType": "_OCFlowDisplayType",
    "TempSensorRange": "_OCTempSensorRange",
    "TempSensorResolution": "_OCTempSensorResolution",
    "TempSensorAccuracy": "_OCTempSensorAccuracy",
    "Connection1":"_Connection1",
    "Connection2":"_Connection2",    
    "ElectricalProtection": "_PROT_TYPE",
    "GasGroup": "_GASGROUP",
    "TempCategory": "_TEMP_CLASS",
    "IPRating": "_IP_RATING",   
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

    