from PointsMaster import PointsMaster
import argparse
import sys
import openpyxl

POINTS_LIST_FILE = "archive/mstm2.xlsx"
SHEET_NAME = 'DNP3.0 Points List'
OUTPUT_FILE = "archive/mtstorm_errors.txt"


def main():

    pm = PointsMaster(POINTS_LIST_FILE, SHEET_NAME, OUTPUT_FILE)
    #pm.verify_device_types_have_all_points()
    pm.validate_point_names()
    pm.validate_point_length()
    pm.validate_underscore_usage()
    pm.validate_point_descriptions()
    pm.validate_all_placeholders_are_removed()
    pm.check_for_duplicates()
    pm.verify_device_ids()
    pm.validate_dnp_indexes_not_reused()

    print("{} Errors Found".format(pm.get_error_count()))

if __name__ == '__main__':
    main()
