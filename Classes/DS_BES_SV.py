# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_SV(Instrument):
    # PROCESS CONDITIONS
    ProcSizCase = ""
    Fluid = ""
    FluidStatPresent = ""
    RelCapReq = ""
    RelCapSelect = ""
    FluidSgDenMolWt = ""
    SpeHeatRatioCpCv = ""
    Compressibilty = ""
    Viscosity = ""
    CriticPress = ""
    CriticTemp = ""
    VapPress = ""
    LatHeatVap = ""
    BlowDown = ""
    AmbTempReqs = ""
    RelCapReqUnits = ""
    FluidSgDenMolWtUnits = ""
    ViscosityUnits = ""
    CricPressUnits = ""
    CriticTempUnits = ""
    VapPressUnits = ""
    ApplicabCode = ""
    OpTemp = ""
    RelTemp = ""
    DesgnTemp = ""
    OpPress = ""
    DesPress = ""
    RelPress = ""
    AtmPress = ""
    ColdDiffPress = ""
    ConstBackPress = ""
    BtUpBckPress = ""
    TotBckPress = ""
    Setpoint = ""
    PcntAllowOvrPress = ""
    OpTempUnits = ""
    OpPressUnits = ""
    # VALVE BODY / BONNET
    ValvNozzleType = ""
    ValvFunct = ""
    ValvTestGauge = ""
    ValvNozzDiscServ = ""
    ValvCapType = ""
    ValvTestGAG = ""
    ValveType = ""
    ValvOperatType = ""
    ValvLeverType = ""
    ValvAsmeSizingCode = ""
    ValvBonnetStyle = ""
    ValvInletConn = ""
    ValvOutletConn = ""
    ValvSeatLeak = ""
    ValvOrifDesig = ""
    ValvOrifAreaCalc = ""
    ValvOrifAreaSelect = ""
    ValveBdyBntMat = ""
    ValvNozzMat = ""
    ValvSoftSeatMat = ""
    ValvDiscPlugMat = ""
    ValvStemMat = ""
    ValvNACEContrct = ""
    ValvSCFMAirAt60DegF = ""
    ValvLbHrSatSteam = ""
    ValvGPMWaterAt70DegF = ""
    ValvGuideRings = ""
    ValvBellMat = ""
    ValvSprngMat = ""
    ValvBckFlowPrevent = ""
    ValvPilotAct = ""
    ValvPilotVent = ""
    ValvPilotType = ""
    ValvPilotFilterMat = ""
    ValvPilotBdyMat = ""
    ValvPilotTrimMat = ""
    ValvFieldTstConn = ""
    ValvPilotDePressOption = ""
    ValvPilotDiffPressSwtch = ""
    ValvCastMat = ""
    ValvRadiography = ""
    ValvMAGParticle = ""
    ValvLPI = ""
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
    "ProcSizCase": "_PCProcSizCase",
    "Fluid": "_PCFluid",
    "FluidStatPresent": "_PCFluidStatPresent",
    "RelCapReq": "_PCRelCapReq",
    "RelCapSelect": "_PCRelCapSelect",
    "FluidSgDenMolWt": "_PCFluidSgDenMolWt",
    "SpeHeatRatioCpCv": "_PCSpeHeatRatioCpCv",
    "Compressibilty": "_PCCompressibilty",
    "Viscosity": "_PCViscosity",
    "CriticPress": "_PCCriticPress",
    "CriticTemp": "_PCCriticTemp",
    "VapPress": "_PCVapPress",
    "LatHeatVap": "_PCLatHeatVap",
    "BlowDown": "_PCBlowDown",
    "AmbTempReqs": "_PCAmbTempReqs",
    "RelCapReqUnits": "_PCRelCapReqUnits",
    "FluidSgDenMolWtUnits": "_PCFluidSgDenMolWtUnits",
    "ViscosityUnits": "_PCViscosityUnits",
    "CricPressUnits": "_PCCricPressUnits",
    "CriticTempUnits": "_PCCriticTempUnits",
    "VapPressUnits": "_PCPVVapPressUnits",
    "ApplicabCode": "_PCApplicabCode",
    "OpTemp": "_PCOpTemp",
    "RelTemp": "_PCRelTemp",
    "DesgnTemp": "_PCDesgnTemp",
    "OpPress": "_PCOpPress",
    "DesPress": "_PCDesPress",
    "RelPress": "_PCRelPress",
    "AtmPress": "_PCAtmPress",
    "ColdDiffPress": "_PCColdDiffPress",
    "ConstBackPress": "_PCConstBackPress",
    "BtUpBckPress": "_PCBtUpBckPress",
    "TotBckPress": "_PCTotBckPress",
    "Setpoint": "_PCSETPOINT",
    "PcntAllowOvrPress": "_PCPcntAllowOvrPress",
    "OpTempUnits": "_PCOpTempUnits",
    "OpPressUnits": "_PCOpPressUnits",
    "ValvNozzleType": "_ValvNozzleType",
    "ValvFunct": "_ValvFunct",
    "ValvTestGauge": "_ValvTestGauge",
    "ValvNozzDiscServ": "_ValvNozzDiscServ",
    "ValvCapType": "_ValvCapType",
    "ValvTestGAG": "_ValvTestGAG",
    "ValveType": "_ValveType",
    "ValvOperatType": "_ValvOperatType",
    "ValvLeverType": "_ValvLeverType",
    "ValvAsmeSizingCode": "_ValvAsmeSizingCode",
    "ValvBonnetStyle": "_ValvBonnetStyle",
    "ValvInletConn": "_ValvInletConn",
    "ValvOutletConn": "_ValvOutletConn",
    "ValvSeatLeak": "_ValvSeatLeak",
    "ValvOrifDesig": "_ValvOrifDesig",
    "ValvOrifAreaCalc": "_ValvOrifAreaCalc",
    "ValvOrifAreaSelect": "_ValvOrifAreaSelect",
    "ValveBdyBntMat": "_ValveBdyBntMat",
    "ValvNozzMat": "_ValvNozzMat",
    "ValvSoftSeatMat": "_ValvSoftSeatMat",
    "ValvDiscPlugMat": "_ValvDiscPlugMat",
    "ValvStemMat": "_ValvStemMat",
    "ValvNACEContrct": "_ValvNACEContrct",
    "ValvSCFMAirAt60DegF": "_ValvSCFMAirAt60DegF",
    "ValvLbHrSatSteam": "_ValvLbHrSatSteam",
    "ValvGPMWaterAt70DegF": "_ValvGPMWaterAt70DegF",
    "ValvGuideRings": "_ValvGuideRings",
    "ValvBellMat": "_ValvBellMat",
    "ValvSprngMat": "_ValvSprngMat",
    "ValvBckFlowPrevent": "_ValvBckFlowPrevent",
    "ValvPilotAct": "_ValvPilotAct",
    "ValvPilotVent": "_ValvPilotVent",
    "ValvPilotType": "_ValvPilotType",
    "ValvPilotFilterMat": "_ValvPilotFilterMat",
    "ValvPilotBdyMat": "_ValvPilotBdyMat",
    "ValvPilotTrimMat": "_ValvPilotTrimMat",
    "ValvFieldTstConn": "_ValvFieldTstConn",
    "ValvPilotDePressOption": "_ValvPilotDePressOption",
    "ValvPilotDiffPressSwtch": "_ValvPilotDiffPressSwtch",
    "ValvCastMat": "_ValvCastMat",
    "ValvRadiography": "_ValvRadiography",
    "ValvMAGParticle": "_ValvMAGParticle",
    "ValvLPI": "_ValvLPI",
    "Note_Field1": "_Notes1",
    "Note_Field2": "_Notes2",
    "Note_Field3": "_Notes3",
    "Note_Field4": "_Notes4",
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



