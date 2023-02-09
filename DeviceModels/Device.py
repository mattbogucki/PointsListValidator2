import copy
from typing import Dict
from typing import List
from typing import Tuple

from DeviceModels.DeviceType import DeviceType


class Device(object):

    ATTRIBUTE = 0
    UNITS = 1
    POINT_TYPE = 2
    DESCRIPTION = 3

    # Schema = [(Attribute Name, Units, Point Type, Description),...]
    def __init__(self, schema: List[Tuple[str, str, str, str]], name: str):
        self._schema = schema
        self._name = name
        self._attribute_units_dictionary = None
        self._attribute_descriptions_dictionary = None
        self._descriptions_set = None
        self._attributes_set = None
        self._required_attributes = None

    # Attribute -> Units
    def get_attribute_units_dictionary(self) -> Dict[str, str]:
        if self._attribute_units_dictionary:
            return copy.deepcopy(self._attribute_units_dictionary)

        new_dict = {}
        for row in self._schema:
            attribute = row[Device.ATTRIBUTE].strip()
            units = row[Device.UNITS].strip() if row[Device.UNITS] else ""
            new_dict[attribute] = units

        self._attribute_units_dictionary = new_dict
        return copy.deepcopy(new_dict)

    # Attribute -> Description
    def get_attribute_descriptions_dictionary(self) -> Dict[str, str]:
        if self._attribute_descriptions_dictionary:
            return copy.deepcopy(self._attribute_descriptions_dictionary)

        new_dict = {}
        for row in self._schema:
            attribute = row[Device.ATTRIBUTE].strip()
            description = row[Device.DESCRIPTION].strip() if row[Device.DESCRIPTION] else ""
            new_dict[attribute] = description

        self._attribute_descriptions_dictionary = new_dict
        return copy.deepcopy(new_dict)

    def get_list_of_required_attributes(self) -> [str]:
        if self._required_attributes:
            return list(self._required_attributes)

        new_list = []
        for row in self._schema:
            attribute = row[Device.ATTRIBUTE].strip()
            new_list.append(attribute)

        self._required_attributes = new_list
        return list(new_list)

    @staticmethod
    def get_regex_pattern_for(device_type) -> str:
        if device_type == DeviceType.METER.value:
            return "^[A-Z\-\d]* HS$|^[A-Z\-\d]* LS$"
        elif device_type == DeviceType.BREAKER_RELAY.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.BUS_RELAY.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.TRANSFORMER_RELAY.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.LINE_RELAY.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.REACTIVE_POWER_RELAY.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.MET_STATION.value:
            return "^MET\d{1,3}$"
        elif device_type == DeviceType.PPC.value:
            return "PPC"
        elif device_type == DeviceType.TOP.value:
            return "TOP"
        elif device_type == DeviceType.MEDIUM_HIGH_VOLTAGE_XFMR.value:
            return "^XFMR(?:\d{3}|\d{3}-.*)$"
        elif device_type == DeviceType.INVERTER.value:
            return "^INV\d{3}$"
        elif device_type == DeviceType.TRACKER_CONTROLLER.value:
            return "^TC\d{3}\-\d{2}_M\d{3}$"
        elif device_type == DeviceType.SOILING_STATION.value:
            return "^SOIL\d{3}$"
        elif device_type == DeviceType.DISCRETE_IO.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.ANALOG_IO.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.ANNUNCIATOR.value:
            return "^[A-Z\-\d]*$"
        elif device_type == DeviceType.INVERTER_MODULE.value:
            return "^INV\d{3}([A-Z]|-\d{2,3})$"
        else:
            return ""




