# -*- coding: utf-8 -*-
from Classes.Instrument import Instrument

class DS_BES_RO(Instrument):
    # PROCESS CONDITIONS
    FluidType = ""
    FluidStatePresent = ""
    Compressibility = ""
    OperatingTemperature = ""
    OperatingPressure = ""
    VaporPressure = ""
    BasePressure = ""
    BaseTemperature = ""
    TempUnits = ""
    PressureUnits = ""
    VaporPressureUnits = ""
    MinLiquidFlow = ""
    NormLiquidFlow = ""
    MaxLiquidFlow = ""
    LiquidFlowUnits = ""
    MinVaporFlow = ""
    NormVaporFlow = ""
    MaxVaporFlow = ""
    VaporFlowUnits = ""
    MinDifferentialFlow = ""
    NormDifferentialFlow = ""
    MaxDifferentialFlow = ""
    DifferentialFlowUnits = ""
    OperatingLiquidDensity = ""
    LiquidDensityUnits = ""
    OperatingViscosity = ""
    ViscosityUnits = ""
    OperatingVaporDensity = ""
    VaporDensityUnits = ""
    SpecificHeatRatio = ""
    SolidsPercentage = ""
    SteamQuality = ""
    PipeMaterial = ""
    FlangeMaterial = ""
    FlangeRating = ""
    # ORIFICE PLATES
    OrificePlateType = ""
    OPCalculationStandard = ""
    OrificePlateMaterial = ""
    OrificePlateInletEdgeStyle = ""
    OrificePlateOutletEdgeStyle = ""
    OrificePlateVentDrainHole = ""
    OrificePlateBetaRatio = ""
    OrificePlate60FBore = ""
    OrificePlateLocation = ""
    OrificePlateOD = ""
    OrificePlateThickness = ""
    # VENTURI / WEDGE METER
    VenturiWedgeType = ""
    VWCalculationStandard = ""
    VenturiWedgeTapType = ""
    VenturiWedgeTapConnection = ""
    VenturiWedgeTapSize = ""
    VenturiWedgeTapRating = ""
    VenturiWedgeBetaRatio = ""
    VenturiWedgeBore = ""
    VenturiWedgeTapOrientation = ""
    VenturiWedgeUpstreamTap = ""
    VenturiWedgeDownstreamTap = ""
    VenturiWedgeTubeEnd = ""
    VenturiWedgeTubeMaterial = ""
    VenturiWedgeSchedule = ""
    VenturiWedgeTubeLining = ""
    VenturiWedgeConvAngle = ""
    VenturiWedgeDivAngle = ""
    # ORIFICE FLANGE SET
    OrificeFlangeSupplier = ""
    OrificeFlangeTapType = ""
    OrificeFlangeTapConnection = ""
    OrificeFlangeTapSize = ""
    OrificeFlangeTapRating = ""
    OrificeFlangeTapOrientation = ""
    OrificeFlangePipeEndConnection = ""
    OrificeFlangeOtherSpecify = ""
    OrificeFlangeHoldingRIngType = ""
    OrificeFlangeHoldingRingMaterial = ""
    OrificeFlangeMaterial = ""
    OrificeFlangeRating = ""
    OrificeFlangeManufacturer = ""
    OrificeFlangeModel = ""
    # PITOT / ANNUBAR
    PitotAnnubarSensorConfiguration = ""
    PitotAnnubarSensorConnection = ""
    PitotAnnubarSensorMaterial = ""
    PitotAnnubarSensorSize = ""
    PitotAnnubarTransmitConnect = ""
    PitotAnnubarValveType = ""
    PitotAnnubarConnection = ""
    PitotAnnubarPipingOrientation = ""
    # FABRICATED METER RUN
    FMRCalculationStandard = ""
    FMRPipeLining = ""
    FMRTapType = ""
    FMRBetaRation = ""
    FMRTapConnection = ""
    FMRBoreDiameter = ""
    FMRTapSize = ""
    FMRUpstreamLength = ""
    FMRTapRating = ""
    FMRDownstreamLength = ""
    FMRTapOrientation = ""
    FMRGasketStyle = ""
    FMREndConnection = ""
    FMRGasketMaterial = ""
    FMRPipeMaterial = ""
    FMRPipeSchedule = ""
    # Flow Element
    FETagNo = ""
    # FLOW NOZZLE
    FNType = ""
    FlowNozzleCalculationStandard = ""
    FlowNozzlePipeEndConnection = ""
    FlowNozzlePipeMaterial = ""
    FlowNozzlePipeSchedule = ""
    FlowNozzleType = ""
    FlowNozzleInnerDiameter = ""
    FlowNozzleMaterial = ""
    FlowNozzleOuterDiameter = ""
    FlowNozzleUpstreamTapConn = ""
    FlowNozzleDownstreamConn = ""
    FlowNozzleTapSize = ""
    FlowNozzleTapRating = ""
    FlowNozzleBetaRatio = ""
    FlowNozzleBoreDiameter = ""
    FlowNozzleTapOrientation = ""
    # NOTES
    Note_Field1 = ""
    Note_Field2 = ""
    Note_Field3 = ""
    Remarks = ""
    InternalFieldsDict = {
    "FluidType": "_PCFluid",
    "FluidStatePresent": "_PCFluidStatePresent",
    "Compressibility": "_PCCompressibility",
    "OperatingTemperature": "_PCTempAtOp",
    "OperatingPressure": "_PCPressAtOp",
    "VaporPressure": "_PCVapPress",
    "BasePressure": "_PCBasePress",
    "BaseTemperature": "_PCBaseTemp",
    "TempUnits": "_PCTempUnits",
    "PressureUnits": "_PCPressUnits",
    "VaporPressureUnits": "_PCVapPressUnits",
    "MinLiquidFlow": "_PCLiqFlowMin",
    "NormLiquidFlow": "_PCLiqFlowNorm",
    "MaxLiquidFlow": "_PCLiqFlowMax",
    "LiquidFlowUnits": "_PCLiqFlowUnits",
    "MinVaporFlow": "_PCVapFlowMin",
    "NormVaporFlow": "_PCVaporFlowNorm",
    "MaxVaporFlow": "_PCVapFlowMax",
    "VaporFlowUnits": "_PCVapFlowUnits",
    "MinDifferentialFlow": "_PCDiffFlowMin",
    "NormDifferentialFlow": "_PCDiffFlowNorm",
    "MaxDifferentialFlow": "_PCDiffFlowMax",
    "DifferentialFlowUnits": "_PCDiffFlowUnits",
    "OperatingLiquidDensity": "_PCLiqSGDensityAtOp",
    "LiquidDensityUnits": "_PCLiqSGDensityUnits",
    "OperatingViscosity": "_PCViscosityAtOperating",
    "ViscosityUnits": "_PCViscosityUnits",
    "OperatingVaporDensity": "_PCVapSGDensityAtOp",
    "VaporDensityUnits": "_PCVapSGDensityUnits",
    "SpecificHeatRatio": "_PCSpecHeatRatio",
    "SolidsPercentage": "_PCPercentSoilds",
    "SteamQuality": "_PCSteamPercentQualityOrSuperHeat",
    "PipeMaterial": "_PCPipeMaterial",
    "FlangeMaterial": "_PCFlngMaterial",
    "FlangeRating": "_PCFlngRating",
    "OrificePlateType": "_OPType",
    "OPCalculationStandard": "_OPCalcStandard",
    "OrificePlateMaterial": "_OPMaterial",
    "OrificePlateInletEdgeStyle": "_OPInletEdgeStyle",
    "OrificePlateOutletEdgeStyle": "_OPOutletEdgeStyle",
    "OrificePlateVentDrainHole": "_OPVentDrainHole",
    "OrificePlateBetaRatio": "_OPBetaRatio",
    "OrificePlate60FBore": "_OPBoreAt60F",
    "OrificePlateLocation": "_OPLocation",
    "OrificePlateOD": "_OPPlateOD",
    "OrificePlateThickness": "_OPThickness",
    "VenturiWedgeType": "_VWType",
    "VWCalculationStandard": "_VWCalStandard",
    "VenturiWedgeTapType": "_VWTapType",
    "VenturiWedgeTapConnection": "_VWTapConn",
    "VenturiWedgeTapSize": "_VWTapSize",
    "VenturiWedgeTapRating": "_VWTapRating",
    "VenturiWedgeBetaRatio": "_VWBetaRatio",
    "VenturiWedgeBore": "_VWBoreDia",
    "VenturiWedgeTapOrientation": "_VWTapOrient",
    "VenturiWedgeUpstreamTap": "_VWUpstreamTapLocation",
    "VenturiWedgeDownstreamTap": "_VWDownstreamTapLoc",
    "VenturiWedgeTubeEnd": "_VWTubeEndConn",
    "VenturiWedgeTubeMaterial": "_VWTubeMaterial",
    "VenturiWedgeSchedule": "_VWSchedule",
    "VenturiWedgeTubeLining": "_VWTubeLining",
    "VenturiWedgeConvAngle": "_VWConvergAngle",
    "VenturiWedgeDivAngle": "_VWTDivergAngle",
    "OrificeFlangeSupplier": "_OFSSuppliedBy",
    "OrificeFlangeTapType": "_OFSTapType",
    "OrificeFlangeTapConnection": "_OFSTapConn",
    "OrificeFlangeTapSize": "_OFSTapSize",
    "OrificeFlangeTapRating": "_OFSTapRating",
    "OrificeFlangeTapOrientation": "_OFSTapOrient",
    "OrificeFlangePipeEndConnection": "_OFSFlngPipeEndConn",
    "OrificeFlangeOtherSpecify": "_OFSOtherSpecify",
    "OrificeFlangeHoldingRIngType": "_OFSHoldRingType",
    "OrificeFlangeHoldingRingMaterial": "_OFSHoldRingMaterial",
    "OrificeFlangeMaterial": "_OFSFlngMaterial",
    "OrificeFlangeRating": "_OFSFlngRating",
    "OrificeFlangeManufacturer": "_OFSManufact",
    "OrificeFlangeModel": "_OFSModel",
    "PitotAnnubarSensorConfiguration": "_PASensConfig",
    "PitotAnnubarSensorConnection": "_PASensorConn",
    "PitotAnnubarSensorMaterial": "_PASensMaterial",
    "PitotAnnubarSensorSize": "_PASensorSize",
    "PitotAnnubarTransmitConnect": "_PAXmtrConn",
    "PitotAnnubarValveType": "_PAValvType",
    "PitotAnnubarConnection": "_Paconn",
    "PitotAnnubarPipingOrientation": "_PAPipingOrient",
    "FMRCalculationStandard": "_FMRCalcStndrd",
    "FMRPipeLining": "_FMRPipeLining",
    "FMRTapType": "_FMRTapType",
    "FMRBetaRation": "_FMRBetaRatio",
    "FMRTapConnection": "_FMRTapConn",
    "FMRBoreDiameter": "_FMRBoreDiam",
    "FMRTapSize": "_FMRTapSize",
    "FMRUpstreamLength": "_FMRUpstreamLen",
    "FMRTapRating": "_FMRTapRating",
    "FMRDownstreamLength": "_FMRDwnStream",
    "FMRTapOrientation": "_FMRTapOrient",
    "FMRGasketStyle": "_FMRGaskStyle",
    "FMREndConnection": "_FMRMeterRunEndConn",
    "FMRGasketMaterial": "_FMRGaskMat",
    "FMRPipeMaterial": "_FMRPipeMatrerial",
    "FMRPipeSchedule": "_FMRPipeSch",
    "FETagNo": "FETagNo",
    "FNType": "_FNType",
    "FlowNozzleCalculationStandard": "_FNCalcStandard",
    "FlowNozzlePipeEndConnection": "_FNFlngPipeEndConn",
    "FlowNozzlePipeMaterial": "_FNPipeMat",
    "FlowNozzlePipeSchedule": "_FNSch",
    "FlowNozzleType": "_FNNozzleType",
    "FlowNozzleInnerDiameter": "_FNNozzleID",
    "FlowNozzleMaterial": "_FNNozzleMat",
    "FlowNozzleOuterDiameter": "_FNNozzOD",
    "FlowNozzleUpstreamTapConn": "_FNUpstreamTapConn",
    "FlowNozzleDownstreamConn": "_FNDwnstreamTapConn",
    "FlowNozzleTapSize": "_FNTapSize",
    "FlowNozzleTapRating": "_FNTapRating",
    "FlowNozzleBetaRatio": "_FNBetaRatio",
    "FlowNozzleBoreDiameter": "_FNBoreDiam",
    "FlowNozzleTapOrientation": "_FNTapOrient",
    "Note_Field1": "_Note1",
    "Note_Field2": "_Note2",
    "Note_Field3": "_Note3",
    "Remarks": "_Remarks"
    }


    

