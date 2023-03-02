import collections
import re

from PointModels.AccumulatorPoint import AccumulatorPoint
from PointModels.AnalogInputPoint import AnalogInputPoint
from PointModels.AnalogOutputPoint import AnalogOutputPoint
from Misc.Availibility import Availability
from PointModels.BinaryInputPoint import BinaryInputPoint
from PointModels.BinaryOutputPoint import BinaryOutputPoint
from DeviceModels.Device import Device
from DeviceModels.DeviceManager import DeviceManager
from DeviceModels.DeviceType import DeviceType
from Misc.Logger import Logger
from PointsListModel.PointsList import PointsList


class Validator(object):
    def __init__(self, points_list: PointsList, standards_file: str, log_file: str):
        self._points_list = points_list
        self._standards_file = standards_file
        self._logger = Logger(log_file)
        self._error_count = 0
        self._device_manager = DeviceManager(standards_file)

    def _verify_reactive_power_points_are_marked_available(self):
        self._logger.log_info("Verifying Reactive Power Points are marked Available")
        for point in self._points_list.get_analog_input_points():
            row = point.get_row()
            point_name = point.get_point_name()
            availability = point.get_availability()
            device_type = point.get_device_type()
            if device_type == DeviceType.REACTIVE_POWER_RELAY.value:
                if 'MVAR Contribution' in point_name and availability != Availability.REQUESTED_AVAILABLE.value:
                    error_msg = "Row {} Var Contribution cannot be marked as unavailable. " \
                                "SEL 487V has this point, otherwise it needs to be calculated. " \
                                "VAR = VAR Nameplate(Voltage/Nameplate Voltage)^2".format(row)
                    self._logger.log_error(error_msg)
        self._logger.log_endline()

    def _verify_suggested_names_for_not_requested_points_follow_pascal_case(self):
        self._logger.log_info("Verifying Not Requested-Available points are Pascal Case with Spaces")
        for point in self._points_list.get_all_points():
            row = point.get_row()
            point_name = point.get_point_name()
            availability = point.get_availability()
            if availability == Availability.NOT_REQUESTED_AVAILABLE.value:
                attribute_name = point_name.split("_")[-1]
                attribute_words = attribute_name.split(" ")
                if self.__has_invalid_casing(attribute_words):
                    error_msg = "Row {} Error - {} casing invalid. Not Requested, Available Point " \
                                "names must be Pascal case with spaces between words. " \
                                "i.e. CSCSB1 Switch Control".format(row, attribute_name)
                    self._logger.log_error(error_msg)
        self._logger.log_endline()

    def __has_invalid_casing(self, words) -> bool:
        is_all_caps = True
        for word in words:
            for i in range(len(word)):
                if i == 0 and word[i].islower():
                    return True
                elif i != 0 and word[i].islower():
                    is_all_caps = False
        return True if is_all_caps else False

    def _validate_only_available_points_have_dnp_indexes(self):
        self._logger.log_info("Validating that only Available points have DNP Indexes")
        for point in self._points_list.get_all_points():
            row = point.get_row()
            dnp_index = point.get_dnp_index()
            availability = point.get_availability()
            if dnp_index.isdigit() and availability == Availability.REQUESTED_NOT_AVAILABLE.value:
                error_msg = "Row {} should not get a dnp index since it is not available/applicable".format(row)
                self._logger.log_error(error_msg)
        self._logger.log_endline()

    def _find_obvious_state_table_errors(self):
        self._logger.log_info("Auditing Binary Input State Tables")
        for point in self._points_list.get_binary_input_points():
            row = point.get_row()
            point_name = point.get_point_name().strip().lower()
            state_table = point.get_state_table().strip().lower()
            if "remote" in point_name and "local" in point_name and "remote" not in state_table:
                self._logger.log_error("Row {} 0/1 State incorrect, 0=Remote, 1=Local or 1=Remote, 0=Local".format(row))
            if "alarm" in point_name and "alarm" not in state_table:
                self._logger.log_error("Row {} is an alarm point, 0/1 state should be 0=Normal, 1=Alarm".format(row))
            if "breaker status" in point_name:
                if "open" not in state_table or "closed" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, needs to be 0=Open, 1=Closed".format(row))
            if "switch status" in point_name:
                if "open" not in state_table or "closed" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, needs to be 0=Open, 1=Closed".format(row))
            if "breaker fail initiate" in point_name:
                if "normal" not in state_table or "bfi" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, needs to be 0=Normal, 1=BFI".format(row))
            if "line comm cutout" in point_name:
                if "normal" not in state_table or "cutout" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, needs to be 0=Normal, 1=Cutout".format(row))
            if "communication status" in point_name:
                if "offline" not in state_table or "online" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, needs to be 0=Offline, 1=Online".format(row))
            if "relay enabled" in point_name:
                if "disabled" not in state_table or "enabled" not in state_table:
                    self._logger.log_error("Row {} 0/1 State is incorrect, must be 0=Disabled, 1=Enabled".format(row))
            if state_table == 'none':
                self._logger.log_error("Row {} 0/1 State table is missing".format(row))
        self._logger.log_endline()

    def _validate_dnp_indexes_not_reused(self):
        self._logger.log_info("Validating that DNP Indexes are not re-used")

        analog_input_index_dictionary = {}
        binary_input_index_dictionary = {}
        counters_index_dictionary = {}
        analog_output_index_dictionary = {}
        binary_output_index_dictionary = {}

        for point in self._points_list.get_all_points():
            dnp_index = point.get_dnp_index()
            row = point.get_row()
            duplicate_row = None

            if not dnp_index.isdigit():
                continue    # DNP Indexes must be ints

            if isinstance(point, AnalogInputPoint):
                duplicate_row = analog_input_index_dictionary.get(dnp_index)
                analog_input_index_dictionary[dnp_index] = row

            elif isinstance(point, AnalogOutputPoint):
                duplicate_row = analog_output_index_dictionary.get(dnp_index)
                analog_output_index_dictionary[dnp_index] = row

            elif isinstance(point, BinaryInputPoint):
                duplicate_row = binary_input_index_dictionary.get(dnp_index)
                binary_input_index_dictionary[dnp_index] = row

            elif isinstance(point, BinaryOutputPoint):
                duplicate_row = binary_output_index_dictionary.get(dnp_index)
                binary_output_index_dictionary[dnp_index] = row

            elif isinstance(point, AccumulatorPoint):
                duplicate_row = counters_index_dictionary.get(dnp_index)
                counters_index_dictionary[dnp_index] = row

            if duplicate_row:
                self._logger.log_error("Row {} used the same DNP index as row {}".format(row, duplicate_row))

        self._logger.log_endline()

    def _verify_substation_name_in_points(self):
        self._logger.log_info("Verifying Substation is included in Point Name")
        for point in self._points_list.get_all_points():
            point_name = point.get_point_name()
            row = point.get_row()
            substation = point.get_substation()
            substation_with_underscores = "_" + substation + "_"
            if substation_with_underscores not in point_name:
                error_msg = "Row {} Point Name missing substation in hierarchy, {}".format(row, point_name, substation)
                self._logger.log_error(error_msg)
        self._logger.log_endline()

    def _verify_device_ids_follow_standard(self):
        self._logger.log_info("Verifying Device IDs - Must be All Caps and follow Attachment 3 Schemas")
        for point in self._points_list.get_all_points():
            device_type = point.get_device_type()
            point_name = point.get_point_name()
            device_id = point.get_source_device()
            row = point.get_row()
            pattern = Device.get_regex_pattern_for(device_type)

            # Verify that the device id was used in the point name
            if device_id not in point_name:
                self._logger.log_error("Row {} Device ID {} not used in Point name".format(row, device_id))

            if pattern:
                # Verify that Device ID follows the defined convention
                match = re.search(pattern, device_id)
                if not match:
                    error_msg = "Row {} Device ID {} invalid for Device Type {}".format(row, device_id, device_type)
                    self._logger.log_error(error_msg)

        self._logger.log_endline()

    def _check_for_duplicates(self):
        self._logger.log_info("Checking For Duplicate Points")

        device_dictionary = {}

        for point in self._points_list.get_all_points():
            point_name = point.get_point_name()
            source_device = point.get_source_device()

            if source_device not in device_dictionary:
                device_dictionary[source_device] = [point_name]
            else:
                device_dictionary[source_device].append(point_name)

        for source_dev, points in device_dictionary.items():
            duplicates = [item for item, count in collections.Counter(points).items() if count > 1]
            for duplicate in duplicates:
                self._logger.log_error("Duplicate point for {}, {}".format(source_dev, duplicate))

        self._logger.log_endline()

    def _verify_devices_have_all_points(self):
        self._logger.log_info("Verifying that each DeviceID has all required attributes")

        device_dictionary = {}
        allowed_device_types = set(item.value for item in DeviceType)

        for point in self._points_list.get_all_points():
            device_type = point.get_device_type()
            if device_type not in allowed_device_types:
                continue  # Other method will flag this error, can't move forward with this point
            source_device = point.get_source_device()
            point_name = point.get_point_name()
            if source_device not in device_dictionary:
                device = self._device_manager.create_device(device_type)
                device_dictionary[source_device] = device.get_list_of_required_attributes()

            dictionary_pts = device_dictionary.get(source_device)
            for pt in dictionary_pts.copy():  # Have to copy or RuntimeError: Set changed size during iteration
                if '#' in pt:
                    string_split = pt.split("#")
                    part1 = string_split[0]
                    part2 = string_split[1]
                    if part1 in point_name and part2 in point_name:
                        dictionary_pts.remove(pt)
                elif pt in point_name:
                    dictionary_pts.remove(pt)

        for source_dev, remaining_pts in device_dictionary.items():
            for remaining_pt in remaining_pts:
                self._logger.log_error("Missing point for {}, {}".format(source_dev, remaining_pt))

        self._logger.log_endline()

    def _verify_units_are_correct(self):
        self._logger.log_info("Verifying Units are correct for all requested points")

        analog_input_points = self._points_list.get_analog_input_points()
        analog_output_points = self._points_list.get_analog_output_points()
        accumulator_points = self._points_list.get_accumulator_points()
        points_with_units = analog_input_points + analog_output_points + accumulator_points
        all_device_types = set(item.value for item in DeviceType)

        for point in points_with_units:
            device_type = point.get_device_type()
            units = point.get_units()
            row = point.get_row()
            point_name = point.get_point_name()
            point_attribute = point_name.split("_")[-1]
            if units == "None":
                self._logger.log_error("Row {} - Missing Units".format(row))
                continue
            #  Other check will flag invalid Device Types, skip over here
            if device_type not in all_device_types:
                continue

            # Device Type -> Dict[Attribute: Unit]
            device_type_dict = {}
            attribute_dictionary = device_type_dict.get(device_type)
            if attribute_dictionary:
                attribute_to_units_dictionary = attribute_dictionary
            else:
                new_device_type = self._device_manager.create_device(device_type)
                attribute_to_units_dictionary = new_device_type.get_attribute_units_dictionary()
                device_type_dict[device_type] = attribute_to_units_dictionary

            correct_units = attribute_to_units_dictionary.get(point_attribute)
            if correct_units and units != correct_units:
                self._logger.log_error("Row {} - {} - Invalid Unit, should be {}".format(row, units, correct_units))
        self._logger.log_endline()

    def _validate_point_descriptions(self):
        self._logger.log_info("Validating point descriptions match A11 for all requested points")
        dm = DeviceManager(self._standards_file)
        all_device_types = set(item.value for item in DeviceType)

        for point in self._points_list.get_all_points():
            device_type = point.get_device_type()
            description = point.get_description()
            row = point.get_row()
            point_name = point.get_point_name()
            point_attribute = point_name.split("_")[-1]
            if description == "None":
                self._logger.log_error("Row {} - Missing Description".format(row))
                continue
            #  Other check will flag invalid Device Types, skip over here
            if device_type not in all_device_types:
                continue

            # Device Type -> Dict[Attribute: Description]
            device_type_dict = {}
            attribute_dictionary = device_type_dict.get(device_type)
            if attribute_dictionary:
                attribute_to_descriptions_dictionary = attribute_dictionary
            else:
                new_device_type = dm.create_device(device_type)
                attribute_to_descriptions_dictionary = new_device_type.get_attribute_descriptions_dictionary()
                device_type_dict[device_type] = attribute_to_descriptions_dictionary

            correct_description = attribute_to_descriptions_dictionary.get(point_attribute)
            if correct_description and description != correct_description:
                error_msg = "Row {} - {} Description should be {}".format(row, description, correct_description)
                self._logger.log_error(error_msg)

        self._logger.log_endline()

    def _validate_all_placeholders_are_removed(self):
        self._logger.log_info("Verifying all Placeholders (#) are removed/defined for available points")
        for point in self._points_list.get_all_points():
            row = point.get_row()
            point_name = point.get_point_name()
            availability = point.get_availability()
            if '#' in point_name and availability != Availability.REQUESTED_NOT_AVAILABLE.value and \
                    availability != Availability.REQUESTED_NOT_APPLICABLE.value:
                self._logger.log_error("Row {} - Placeholder left in point {}".format(row, point_name))

        self._logger.log_endline()

    def _verify_engineering_units_defined_for_analogs(self):
        self._logger.log_info("Verifying valid EGU Min and Max Units are defined for all Analog Inputs")
        for point in self._points_list.get_analog_input_points():
            row = point.get_row()
            availability = point.get_availability()
            if availability == Availability.REQUESTED_AVAILABLE.value or \
                    availability == Availability.NOT_REQUESTED_AVAILABLE.value:
                egu_min = point.get_egu_min()
                egu_max = point.get_egu_max()
                if egu_min == 'None' or egu_max == 'None':
                    self._logger.log_error("Row {} - Missing Engineering Unit Min/Max".format(row))
                    continue
                try:
                    egu_min = float(egu_min)
                    egu_max = float(egu_max)
                except:
                    self._logger.log_error("Row {} - Invalid Engineering Units, Must be numeric".format(row))
                    continue

                if egu_min > egu_max:
                    self._logger.log_error("Row {} - Invalid Engineering Units, Min must be less than Max".format(row))

        self._logger.log_endline()

    def _verify_availability_entry_is_valid(self):
        self._logger.log_info("Verifying Availability field is set to a valid state")
        for point in self._points_list.get_all_points():
            row = point.get_row()
            availability = point.get_availability()
            allowed_availabilities = set(item.value for item in Availability)
            if availability not in allowed_availabilities:
                error_msg = "Row {} - {} - Invalid Availability. Valid options are Requested-Not Available, " \
                            "Requested-Available, Requested-Not Applicable, Not Requested-Available"\
                            .format(row, availability)

                self._logger.log_error(error_msg)

        self._logger.log_endline()

    def _validate_underscore_usage(self):
        self._logger.log_info("Validating Underscore Usage")
        for point in self._points_list.get_all_points():
            point_name = point.get_point_name()
            row = point.get_row()
            device_type = point.get_device_type()
            number_of_underscores = point_name.count("_")
            if device_type == DeviceType.INVERTER_MODULE.value:  # Inverter Module is only device allowed 4
                if number_of_underscores != 3 and number_of_underscores != 4:
                    error_msg1 = "Row {} - {} - has {} underscores when there should be 3 or 4 for Inv Modules per A11"\
                        .format(row, point_name, number_of_underscores)
                    self._logger.log_error(error_msg1)
            elif number_of_underscores != 3:
                error_msg2 = "Row {} - {} - has {} underscores when there should be exactly 3 per A11"\
                    .format(row, point_name, number_of_underscores)
                self._logger.log_error(error_msg2)

        self._logger.log_endline()

    def _validate_point_length(self):
        self._logger.log_info("Verifying All Not Requested-Available points are less than 60 chars")
        for point in self._points_list.get_all_points():
            point_name = point.get_point_name()
            name_length = len(point_name)
            row = point.get_row()
            availability = point.get_availability()
            if availability != Availability.NOT_REQUESTED_AVAILABLE.value:
                continue
            if name_length > 60:
                error_msg = "Row {} - Point Exceeds 60 char limit by {} chars".format(row, name_length - 60)
                self._logger.log_error(error_msg)

        self._logger.log_endline()

    def _validate_device_types(self):
        self._logger.log_info("Verifying Device Types are valid")

        all_device_types = set(item.value for item in DeviceType)
        for point in self._points_list.get_all_points():
            row = point.get_row()
            device_type = point.get_device_type()
            if device_type not in all_device_types:
                self._logger.log_error("Row {} - Source Device Type {} invalid".format(row, device_type))

        self._logger.log_endline()

    def _validate_point_names(self):
        self._logger.log_info("Validating Point Names Against A11")

        # Breaker Relay -> [attribute 1, attribute 2, ... attribute n]
        device_type_to_attributes_dict = {}

        all_device_types = set(item.value for item in DeviceType)

        for point in self._points_list.get_all_points():
            point_name = point.get_point_name()
            row = point.get_row()
            is_available = point.get_availability()
            device_type = point.get_device_type()

            if device_type not in all_device_types:
                continue    # Skip since we have nothing to compare against

            if is_available == Availability.NOT_REQUESTED_AVAILABLE.value:
                continue    # Attributes we didn't explicitly request will have made up names

            if device_type_to_attributes_dict.get(device_type):
                valid_attributes = device_type_to_attributes_dict.get(device_type)
            else:
                new_device_type = self._device_manager.create_device(device_type)
                valid_attributes = new_device_type.get_list_of_required_attributes()

            if not self.__point_is_valid(point_name, valid_attributes):
                self._logger.log_error("Row {} - {} - Invalid Point Name".format(row, point_name))

        self._logger.log_endline()

    def __point_is_valid(self, name: str, valid_attributes: [str]) -> bool:
        for valid_name in valid_attributes:
            if valid_name in name:
                return True
            if '#' in valid_name:
                string_split = valid_name.split("#")
                part1 = string_split[0]
                part2 = string_split[1]
                if part1 in name and part2 in name:
                    return True
        return False

    def _validate_pi_ip_address_is_set(self):
        pi_ip_address = self._points_list.get_pi_ip_address()
        if pi_ip_address is None:
            error_msg = "IP Address for Pi Connection must be set in cell B2"
            self._logger.log_error(error_msg)
        else:
            pi_match = re.search("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", pi_ip_address)
            if not pi_match:
                error_msg = "Invalid IP Address set in cell B2"
                self._logger.log_error(error_msg)

    def _validate_gms_ip_address_is_set(self):
        gms_ip_address = self._points_list.get_gms_ip_address()
        if gms_ip_address is None:
            error_msg = "IP Address for GMS Connection must be set in cell E2"
            self._logger.log_error(error_msg)
        else:
            gms_match = re.search("\d{1,3}\.\d{1,3}\.\d{1,3}\.\d{1,3}", gms_ip_address)
            if not gms_match:
                error_msg = "Invalid IP Address set in cell E2"
                self._logger.log_error(error_msg)

    def validate_points_list(self):
        self._verify_reactive_power_points_are_marked_available()
        self._verify_suggested_names_for_not_requested_points_follow_pascal_case()
        self._validate_only_available_points_have_dnp_indexes()
        self._find_obvious_state_table_errors()
        self._validate_dnp_indexes_not_reused()
        self._verify_substation_name_in_points()
        self._verify_device_ids_follow_standard()
        self._check_for_duplicates()
        self._verify_devices_have_all_points()
        self._verify_units_are_correct()
        self._validate_point_descriptions()
        self._validate_all_placeholders_are_removed()
        self._verify_engineering_units_defined_for_analogs()
        self._verify_availability_entry_is_valid()
        self._validate_underscore_usage()
        self._validate_point_length()
        self._validate_device_types()
        self._validate_point_names()
        self._validate_pi_ip_address_is_set()
        self._validate_gms_ip_address_is_set()
        self._logger.print_error_count()

