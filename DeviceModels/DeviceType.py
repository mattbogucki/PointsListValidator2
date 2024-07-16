from enum import Enum


class DeviceType(Enum):
    BREAKER_RELAY = 'Breaker Relay'
    BUS_RELAY = 'Bus Relay'
    TRANSFORMER_RELAY = 'Transformer Relay'
    TRANSFORMER_DEVICE = 'Transformer Device'
    LINE_RELAY = 'Line Relay'
    REACTIVE_POWER_RELAY = 'Reactive Power Relay'
    MEDIUM_HIGH_VOLTAGE_XFMR = 'Medium-High Voltage Transformer'
    TRACKER_CONTROLLER = 'Tracker Controller'
    PPC = 'Generator PPC'
    TOP = 'TOP'
    ANALOG_IO = 'Analog IO (Axion)'
    DISCRETE_IO = 'Discrete IO (DPAC)'
    MET_STATION = 'MET Station'
    INVERTER = 'Inverter'
    INVERTER_MODULE = 'Inverter Module'
    METER = 'Meter'
    SOILING_STATION = 'Soiling Station'
    ANNUNCIATOR = 'Annunciator'
    ERCOT = 'ERCOT'
    POI = 'POI'
    BREAKER_AND_MOD_CONTROLS = 'Breaker and MOD Controls'
    WARTSILA_ESS_UNIT = 'Wartsila ESS Unit'

