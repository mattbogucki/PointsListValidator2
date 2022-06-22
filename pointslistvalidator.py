import argparse
import os
import sys

import openpyxl

from PointsMaster import PointsMaster


def main():

    parser = argparse.ArgumentParser(description='Points List Validator')
    parser.add_argument('FILENAME', action="store", help=".xlsx filename of points list")
    parser.add_argument('TAB', action="store", help='Name of tab with owners points list (Default is DNP3.0 Points List)')
    try:
        cmd_line_args = parser.parse_args()
    except:
        print('Supply the proper arguments')
        sys.exit(-1)

    try:
        test_wb = openpyxl.load_workbook(cmd_line_args.FILENAME, data_only=True)
        test_sheet = test_wb[cmd_line_args.TAB]
    except:
        print('File Name and Sheet not valid')
        sys.exit(-2)

    desktop_logfile = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop\point_list_validator_log.txt')

    pm = PointsMaster(cmd_line_args.FILENAME, cmd_line_args.TAB, desktop_logfile)
    pm.verify_device_types_have_all_points()
    pm.validate_point_names()
    pm.validate_point_length()
    pm.validate_underscore_usage()
    pm.validate_point_descriptions()
    pm.validate_all_placeholders_are_removed()
    pm.check_for_duplicates()
    pm.verify_device_ids()
    pm.validate_dnp_indexes_not_reused()
    pm.verify_substation_name_in_points()
    pm.validate_only_available_points_have_dnp_indexes()
    pm.find_obvious_state_table_errors()
    pm.verify_suggested_names_for_not_requested_points_follow_pascal_case()
    pm.verify_availability_entry_is_valid()
    pm.verify_units_are_correct()

    with open(desktop_logfile, mode="a") as f:
        f.write("{} Errors Found".format(pm.get_error_count()))
        print("{} Errors Found".format(pm.get_error_count()))


if __name__ == '__main__':
    main()