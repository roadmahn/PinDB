# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_Weatherstation(Instrument):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    # WEATHERSTATION = ""
    PowerRequirements = ""
    PowerConsumption = ""
    CurrentConsumption = ""
    ReversePolarity = ""
    WindSpeed = ""
    WindSpeedResponse = ""
    WindSpeedVariables = ""
    WindSpeedAccuracy = ""
    WindSpeedResolution = ""
    WindDirAzimuth = ""
    WindDirResponse = ""
    WindDirVariables = ""
    WindDirAccuracy = ""
    WindDirResolution = ""
    Rainfall = ""
    RainfallCollectingArea = ""
    RainfallResolution = ""
    RainDuration = ""
    RainDurationResolution = ""
    RainIntensity = ""
    RainIntensityRange = ""
    RainIntensityResolution = ""
    Hail = ""
    HailResolution = ""
    HailDuration = ""
    HailDurationResolution = ""
    HailIntensity = ""
    HailIntensityResolution = ""
    PressRange = ""
    PressAccuracy = ""
    PressResolution = ""
    TempRange = ""
    TempAccuracy = ""
    TempResolution = ""
    HumidRange = ""
    HumidAccuracy = ""
    HumidResolution = ""
    Heating = ""
    HeatingVoltage = ""
    HeatingCurrent = ""
    CommHardwareLayer = ""
    CommProtocols = ""
    AISolarRadiation = ""
    AILevelMeasurement = ""
    AIRainGauge = ""
    AITemperature = ""
    AOWindSpeed = ""
    AOWindDirection = ""
    AOLoadImpedance = ""
    HousingIPRating = ""
    StorageTemperature = ""
    OperatingTemperature = ""
    RelativeHumidity = ""
    GenPressure = ""
    GenWind = ""   
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
    "PowerRequirements": "_WSPwrRequirements",
    "PowerConsumption": "_WSPwrConsumption",
    "CurrentConsumption": "_WSCurrentConsumption",
    "ReversePolarity": "_WSReversePolarity",
    "WindSpeed": "_WSWindSpeed",
    "WindSpeedResponse": "_WSWindSpeedResponse",
    "WindSpeedVariables": "_WSWindSpeedVariables",
    "WindSpeedAccuracy": "_WSWindSpeedAccuracy",
    "WindSpeedResolution": "_WSWindSpeedResolution",
    "WindDirAzimuth": "_WSWindDirAzimuth",
    "WindDirResponse": "_WSWindDirResponse",
    "WindDirVariables": "_WSWindDirVariables",
    "WindDirAccuracy": "_WSWindDirAccuracy",
    "WindDirResolution": "_WSWindDirResolution",
    "Rainfall": "_WSRainfall",
    "RainfallCollectingArea": "_WSRainfallCollectingArea",
    "RainfallResolution": "_WSRainfallResolution",
    "RainDuration": "_WSRainDuration",
    "RainDurationResolution": "_WSRainDurationResolution",
    "RainIntensity": "_WSRainIntensity",
    "RainIntensityRange": "_WSRainIntensityRange",
    "RainIntensityResolution": "_WSRainIntensityResolution",
    "Hail": "_WSHail",
    "HailResolution": "_WSHailResolution",
    "HailDuration:": "_WSHailDuration",
    "HailDurationResolution": "_WSHailDurationResolution",
    "HailIntensity": "_WSHailIntensity",
    "HailIntensityResolution": "_WSHailIntensityResolution",
    "PressRange": "_WSPressRange",
    "PressAccuracy": "_WSPressAccuracy",
    "PressResolution": "_WSPressResolution",
    "TempRange": "_WSTempRange",
    "TempAccuracy": "_WSTempAccuracy",
    "TempResolution": "_WSTempResolution",
    "HumidRange": "_WSHumidRange",
    "HumidAccuracy": "_WSHumidAccuracy",
    "HumidResolution": "_WSHumidResolution",
    "Heating": "_WSHeating",
    "HeatingVoltage": "_WSHeatingVoltage",
    "HeatingCurrent": "_WSHeatingCurrent",
    "CommHardwareLayer": "_WSCommHardwareLayer",
    "CommProtocols": "_WSCommProtocols",
    "AISolarRadiation": "_WSAISolarRadiation",
    "AILevelMeasurement": "_WSAILevelMeasurement",
    "AIRainGauge": "_WSAIRainGauge",
    "AITemperature": "_WSAITemperature",
    "AOWindSpeed": "_WSAOWindSpeed",
    "AOWindDirection": "_WSAOWindDirection",
    "AOLoadImpedance": "_WSAOLoadImpedance",
    "HousingIPRating": "_WSGenHousingIP",
    "StorageTemperature": "_WSGenStoreTemp",
    "OperatingTemperature": "_WSGenOperTemp",
    "RelativeHumidity": "_WSGenRelHumid",
    "GenPressure": "_WSGenRelPressure",
    "GenWind": "_WSGenWindSpeed",
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

    