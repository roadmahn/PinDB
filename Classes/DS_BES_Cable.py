# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument
from Classes.Termination import Termination

class DS_BES_Cable(Instrument,Termination):
    # GENERAL
    AreaClassification = ""
    AmbientRelHumidity = ""
    AmbientTemperature = ""
    ActivationSource = ""
    Function = ""
    ConductorMaterial = ""
    InsulationMaterial = ""
    IMColour = ""
    JacketSheathMaterial = ""
    FlameRetardancy = ""
    OilResistant = ""
    CableAssembly = ""
    CrossSection = ""
    InsulationThickness = ""
    SheathThickness = ""
    OverallDiameter = ""
    Weight = ""
    PullingTension = ""
    CoreConfiguration = ""
    Connection1 = ""
    Connection2 = ""
    # ELECTRICAL
    VoltageRating = ""
    Resistance = ""
    Capacitance = ""
    Inductance = ""
    Reactance = ""
    Impedance = ""
    InsResistance = ""
    ChargingCurrent = ""
    DielectricLoss = ""
    ShortCircuitCapacity = ""
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
    "ConductorMaterial": "_CnstrConductorMat",
    "InsulationMaterial": "_CnstrInsulationMat",
    "IMColour": "_CnstrIMColour",
    "JacketSheathMaterial": "_CnstrJacketMat",
    "FlameRetardancy": "_CnstrFlameRetardancy",
    "OilResistant": "_CnstrOilResistant",
    "CableAssembly": "_CnstrCableAssembly",
    "CrossSection": "_CnstrCrossSection",
    "InsulationThickness": "_CnstrInsulationThickness",
    "SheathThickness": "_CnstrSheatThickness",
    "OverallDiameter": "_CnstrOverallDiameter",
    "Weight": "_CnstrWeight",
    "PullingTension": "_CnstrPullingTension",
    "CoreConfiguration": "_CnstrCoreConfiguration",
    "Connection1": "_CnstrConnection1",
    "Connection2": "_CnstrConnection2",
    "VoltageRating": "_ElecVoltage",
    "Resistance": "_ElecResistance",
    "Capacitance": "_ElecCapacitance",
    "Inductance": "_ElecInductance",
    "Reactance": "_ElecReactance",
    "Impedance": "_ElecImpedance",
    "InsResistance": "_ElecInsResistance",
    "ChargingCurrent": "_ElecChargingCurrent",
    "DielectricLoss": "_ElecDielectricLoss",
    "ShortCircuitCapacity": "_ElecShortCircuitCapacity",    
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

    