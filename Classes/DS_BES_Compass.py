# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Compass(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    # COMPASS
    AntennaImpedance = ""
    AtlasLband = ""
    Autonomous = ""
    BaudRates = ""
    ColdStart = ""
    Ports = ""
    TimingOutput = ""
    CorrectionProtocol = ""
    CurrentConsumption = ""
    DataProtocol = ""
    DifferentialOptions = ""
    Dimensions = ""
    EMC = ""
    EventMarkerInput = ""
    GPSSensitivity = ""
    Gyro = ""
    Heading = ""
    HeadingWarning = ""
    HeadingFix = ""
    Heave = ""
    HotStart = ""
    Humidity = ""
    ReacquisitionTime = ""
    LBandChannels = ""
    ChannelSpacing = ""
    Processor = ""
    SatelliteSelection = ""
    Sensitivity = ""
    MaximumAltitude = ""
    MaximumSpeed = ""
    Mounting = ""
    OperatingTemperature = ""
    PitchRoll = ""
    PowerConsumption = ""
    PwrDataConnector = ""
    PowerRequirements = ""
    RateofTurn = ""
    Accuracy = ""
    Channels = ""
    SignalsReceived = ""
    ReceiverType = ""
    ReversePolarity = ""
    SBAS = ""
    SBASTracking = ""
    StatusIndications = ""
    StorageTemperature = ""
    TiltSensors = ""
    UpdateRate = ""
    Vibration = ""
    WarmStart = ""
    Weight = ""
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
    "AntennaImpedance": "_CompassAntennaImpedance",
    "AtlasLband": "_CompassAtlasLBandPosition",
    "Autonomous": "_CompassAutonomousPosition",
    "BaudRates": "_CompassBaudRates",
    "ColdStart": "_CompassColdStart",
    "Ports": "_CompassCommPorts",
    "TimingOutput": "_CompassCommTimingOutput",
    "CorrectionProtocol": "_CompassCorrectionProtocol",
    "CurrentConsumption": "_CompassCurrentConsumption",
    "DataProtocol": "_CompassDataProtocol",
    "DifferentialOptions": "_CompassDiffOptions",
    "Dimensions": "_CompassDimensions",
    "EMC": "_CompassEMC",
    "EventMarkerInput": "_CompassEventMarker",
    "GPSSensitivity": "_CompassGPSSensitivity",
    "Gyro": "_CompassGyro",
    "Heading": "_CompassHeadingPosition",
    "HeadingWarning": "_CompassHeadingWarning",
    "HeadingFix": "_CompassHeadingFix",
    "Heave": "_CompassHeavePosition",
    "HotStart": "_CompassHotStart",
    "Humidity": "_CompassHumidity",
    "ReacquisitionTime": "_CompassLBandAcqTime",
    "LBandChannels": "_CompassLBandChannels",
    "ChannelSpacing": "_CompassLBandChannelSpace",
    "Processor": "_CompassLBandCProcessor",
    "SatelliteSelection": "_CompassLBandSatSelection",
    "Sensitivity": "_CompassLBandSensitivity",
    "MaximumAltitude": "_CompassMaxAltitude",
    "MaximumSpeed": "_CompassMaxSpeed",
    "Mounting": "_CompassMounting",
    "OperatingTemperature": "_CompassOperTemperature",
    "PitchRoll": "_CompassPitchRollPosition",
    "PowerConsumption": "_CompassPwrConsumption",
    "PwrDataConnector": "_CompassPwrDataConnector",
    "PowerRequirements": "_CompassPwrRequirements",
    "RateofTurn": "_CompassRateOfTurn",
    "Accuracy": "_CompassReceiverAccuracy",
    "Channels": "_CompassReceiverChannels",
    "SignalsReceived": "_CompassReceiverSignals",
    "ReceiverType": "_CompassReceiverType",
    "ReversePolarity": "_CompassReversePolarity",
    "SBAS": "_CompassSBASPosition",
    "SBASTracking": "_CompassSBASTracking",
    "StatusIndications": "_CompassStatusIndications",
    "StorageTemperature": "_CompassStoreTemperature",
    "TiltSensors": "_CompassTiltSensors",
    "UpdateRate": "_CompassUpdateRate",
    "Vibration": "_CompassVibration",
    "WarmStart": "_CompassWarmStart",
    "Weight": "_CompassWeight",
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

    