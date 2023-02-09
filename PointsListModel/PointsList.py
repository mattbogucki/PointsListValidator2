import sys

import openpyxl

from PointModels.AnalogInputPoint import *
from PointModels.AnalogOutputPoint import *
from PointModels.BinaryInputPoint import *
from PointModels.BinaryOutputPoint import *
from PointModels.AccumulatorPoint import *


class PointsList(object):
    def __init__(self, points_list_wb: str, wb_sheet: str):
        try:
            self._workbook = openpyxl.load_workbook(points_list_wb, data_only=True)
            self._sheet = self._workbook[wb_sheet]
        except:
            print('File Name and Sheet not valid')
            sys.exit(-2)

        self._start = 20  # Ignore boilerplate
        self._HEADER = 0

        self._SUBSTATION = 0
        self._DEVICE_ID = 1
        self._POINT_NAME = 2
        self._DEVICE_TYPE = 3
        self._DESCRIPTION = 5
        self._DNP_INDEX = 6
        self._UNITS = 7
        self._STATE_TABLE = 7
        self._AVAILABLE = 8
        self._EGU_MIN = 9
        self._ROLLOVER = 9
        self._BINARY_INPUT_FEEDBACK_POINT = 9
        self._ANALOG_INPUT_FEEDBACK_POINT = 9
        self._EGU_MAX = 10

        self._analog_inputs = []
        self._analog_outputs = []
        self._binary_inputs = []
        self._binary_outputs = []
        self._accumulators = []

        # Spreadsheet has a header column that announces when point types change
        self._point_type = None

        for i, row in enumerate(self._sheet):
            point_type = str(row[self._SUBSTATION].value).strip()
            self._point_type = self._get_point_type(point_type)
            substation = str(row[self._SUBSTATION].value).strip()
            device_id = str(row[self._DEVICE_ID].value).strip()
            point_name = str(row[self._POINT_NAME].value).strip()
            device_type = str(row[self._DEVICE_TYPE].value).strip()
            description = str(row[self._DESCRIPTION].value).strip()
            dnp_index = str(row[self._DNP_INDEX].value).strip()
            availability = str(row[self._AVAILABLE].value).strip()

            if i < self._start:
                continue

            if not point_name or point_name == 'None' or point_name == 'Point Name':
                continue

            if self._point_type == 'Digital Inputs':
                state_table = str(row[self._STATE_TABLE].value).strip()
                self._binary_inputs.append(
                    BinaryInputPoint(i + 1, substation, point_name, device_id, device_type, description, dnp_index,
                                     availability, state_table))
            elif self._point_type == 'Digital Outputs':
                state_table = str(row[self._STATE_TABLE].value).strip()
                feedback_index = str(row[self._BINARY_INPUT_FEEDBACK_POINT].value).strip()
                self._binary_outputs.append(
                    BinaryOutputPoint(i + 1, substation, point_name, device_id, device_type, description, dnp_index,
                                      availability, state_table, feedback_index))
            elif self._point_type == 'Analog Inputs':
                egu_min = str(row[self._EGU_MIN].value).strip()
                egu_max = str(row[self._EGU_MAX].value).strip()
                units = str(row[self._UNITS].value).strip()
                self._analog_inputs.append(
                    AnalogInputPoint(i + 1, substation, device_id, point_name, device_type, description, dnp_index,
                                     availability, units, egu_min, egu_max))
            elif self._point_type == 'Analog Outputs':
                feedback_index = str(row[self._ANALOG_INPUT_FEEDBACK_POINT].value).strip()
                units = str(row[self._UNITS].value).strip()
                self._analog_outputs.append(
                    AnalogOutputPoint(i + 1, substation, device_id, point_name, device_type, description, dnp_index,
                                      availability, units, feedback_index))
            elif self._point_type == 'Counters':
                units = str(row[self._UNITS].value).strip()
                rollover_value = str(row[self._ROLLOVER].value).strip()
                self._accumulators.append(
                    AccumulatorPoint(i + 1, substation, device_id, point_name, device_type, description, dnp_index,
                                     availability, units, rollover_value))

    #  Point Types are only announced once
    def _get_point_type(self, entry) -> str:
        if entry == 'Analog Inputs':
            return 'Analog Inputs'
        elif entry == 'Analog Outputs':
            return 'Analog Outputs'
        elif entry == 'Digital Inputs':
            return 'Digital Inputs'
        elif entry == 'Digital Outputs':
            return 'Digital Outputs'
        elif entry == 'Counters':
            return 'Counters'
        else:
            return self._point_type

    def get_binary_input_points(self) -> [BinaryInputPoint]:
        return list(self._binary_inputs)

    def get_binary_output_points(self) -> [BinaryOutputPoint]:
        return list(self._binary_outputs)

    def get_analog_input_points(self) -> [AnalogInputPoint]:
        return list(self._analog_inputs)

    def get_analog_output_points(self) -> [AnalogOutputPoint]:
        return list(self._analog_outputs)

    def get_accumulator_points(self) -> [AccumulatorPoint]:
        return list(self._accumulators)

    def get_all_points(self) -> [Point]:
        return list(self._analog_inputs) + \
               list(self._analog_outputs) + \
               list(self._binary_outputs) + \
               list(self._binary_inputs) + \
               list(self._accumulators)
