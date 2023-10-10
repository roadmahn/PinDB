# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_SOL_V(Instrument):
    # PROCESS CONDITIONS
    AreaClassification = ""
    AmbientTemperatureRequirements = ""
    Fluid1RequiredCv = ""
    Fluid1Fluid = ""
    Fluid1MaxTemp = ""
    Fluid1MaxPress = ""
    Fluid1OpDiffPress = ""
    Fluid1MaxDiffPress = ""
    Fluid1MaxAlowDiffPress = ""
    Fluid1SgDensityMolwgt = ""
    Fluid1Viscosity = ""
    Fluid2RequiredCv = ""
    Fluid2Fluid = ""
    Fluid2MaxTemp = ""
    Fluid2MaxPress = ""
    Fluid2OpDiffPress = ""
    Fluid2MaxDiffPress = ""
    Fluid2MaxAlowDiffPress = ""
    Fluid2SgDensityMolwgt = ""
    Fluid2Viscosity = ""
    TempUnits = ""
    PressUnits = ""
    SgDensityMolwgtUnits = ""
    ViscosityUnits = ""
    # SOLENOID
    SolenoidType = ""
    SolenoidCoil = ""
    SolenoidVoltage = ""
    SolenoidPwrWiring = ""
    SolenoidPilotOperated = ""
    SolenoidCommProtocol = ""
    SolenoidLocation = ""
    SolenoidMount = ""
    SolenoidElectricalProtection = ""
    SolenoidTempCategory = ""
    SolenoidBdySize = ""
    SolenoidGasGroup = ""
    SolenoidIPRating = ""
    SolenoidBdyMat = ""
    SolenoidCoilInsulation = ""
    SolenoidDiscMat = ""
    SolenoidMainValvAct = ""
    SolenoidPackingStyle = ""
    SolenoidManOvrRide = ""
    SolenoidGlandConn = ""
    SolenoidManReset = ""
    SolenoidPortConn = ""
    SolenoidManufact = ""
    SolenoidSeatMat = ""
    SolenoidModelNo = ""
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
    "AreaClassification": "_PCAREA_CLASS_REQ",
    "AmbientTemperatureRequirements": "_PCAmbientTempRequirements",
    "Fluid1RequiredCv": "_PCFluid1RequiredCv",
    "Fluid1Fluid": "_PCFluid1Fluid",
    "Fluid1MaxTemp": "_PCFluid1MaxTemp",
    "Fluid1MaxPress": "_PCFluid1MaxPress",
    "Fluid1OpDiffPress": "_PCFluid1OpDiffPress",
    "Fluid1MaxDiffPress": "_PCFluid1MaxDiffPress",
    "Fluid1MaxAlowDiffPress": "_PCFluid1MaxAlowDiffPress",
    "Fluid1SgDensityMolwgt": "_PCFluid1SgDensityMolwgt",
    "Fluid1Viscosity": "_PCFluid1Viscosity",
    "Fluid2RequiredCv": "_PCFluid2RequiredCv",
    "Fluid2Fluid": "_PCFluid2Fluid",
    "Fluid2MaxTemp": "_PCFluid2MaxTemp",
    "Fluid2MaxPress": "_PCFluid2MaxPress",
    "Fluid2OpDiffPress": "_PCFluid2OpDiffPress",
    "Fluid2MaxDiffPress": "_PCFluid2MaxDiffPress",
    "Fluid2MaxAlowDiffPress": "_PCFluid2MaxAlowDiffPress",
    "Fluid2SgDensityMolwgt": "_PCFluid2SgDensityMolwgt",
    "Fluid2Viscosity": "_PCFluid2Viscosity",
    "TempUnits": "_PCTempUnits",
    "PressUnits": "_PCPressUnits",
    "SgDensityMolwgtUnits": "_PCSgDensityMolwgtUnits",
    "ViscosityUnits": "_PCViscosityUnits",
    "SolenoidType": "_SolenoidType",
    "SolenoidCoil": "_SolenoidCoil",
    "SolenoidVoltage": "_SolenoidVoltage",
    "SolenoidPwrWiring": "_SolenoidPwrWiring",
    "SolenoidPilotOperated": "_SolenoidPilotOperated",
    "SolenoidCommProtocol": "_SolenoidCommProtocol",
    "SolenoidLocation": "_SolenoidLocation",
    "SolenoidMount": "_SolenoidMount",
    "SolenoidElectricalProtection": "_PROT_TYPE",
    "SolenoidTempCategory": "_TEMP_CLASS",
    "SolenoidBdySize": "_SolenoidBdySize",
    "SolenoidGasGroup": "_GASGROUP",
    "SolenoidIPRating": "_IP_RATING",
    "SolenoidBdyMat": "_SolenoidBdyMat",
    "SolenoidCoilInsulation": "_SolenoidCoilInsulation",
    "SolenoidDiscMat": "_SolenoidDiscMat",
    "SolenoidMainValvAct": "_SolenoidMainValvAct",
    "SolenoidPackingStyle": "_SolenoidPackingStyle",
    "SolenoidManOvrRide": "_SolenoidManOvrRide",
    "SolenoidGlandConn": "_SolenoidGlandConn",
    "SolenoidManReset": "_SolenoidManReset",
    "SolenoidPortConn": "_SolenoidPortConn",
    "SolenoidManufact": "_SolenoidManufact",
    "SolenoidSeatMat": "_SolenoidSeatMat",
    "SolenoidModelNo": "_SolenoidModelNo",
    "Connection1": "_Connection1",
    "Connection2": "_Connection2",    
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Note_Field4": "_Notes4",
    "Note_Field5": "_Notes5",
    "Note_Field6": "_Notes6",
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
