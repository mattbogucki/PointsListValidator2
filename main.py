import argparse
import os
import sys

import openpyxl

from PointsListModel.PointsList import PointsList
from ValidatorModel.Validator import Validator
from Misc.Expiration import *


def main():

    if version_is_expired():
        print("PointsListValidator version is outdated, please reach out to Clearway for updated version")
        sys.exit(-1)

    parser = argparse.ArgumentParser(description='Points List Validator')
    parser.add_argument('FILENAME', action="store", help=".xlsx filename of points list")
    sheet_help = 'Name of tab with owners points list (Default is DNP3.0 Points List)'
    parser.add_argument('SHEET', action="store", help=sheet_help)

    try:
        cmd_line_args = parser.parse_args()
    except:
        print('Supply the proper command line arguments, use -help for information on correct parameters')
        sys.exit(-1)

    try:
        test_wb = openpyxl.load_workbook(cmd_line_args.FILENAME, data_only=True)
        test_sheet = test_wb[cmd_line_args.SHEET]
    except:
        print('File Name and Sheet not valid')
        sys.exit(-1)

    print("Loading Points List...")
    points_list = PointsList(cmd_line_args.FILENAME, cmd_line_args.SHEET)
    # points_list = PointsList("TestSpreadsheets\d3.xlsx", "1")
    log_file = os.path.join(os.path.join(os.environ['USERPROFILE']), 'Desktop\point_list_validator_log.txt')
    validator = Validator(points_list, "Attachment 3-Attribute Names & Engineering Units.xlsx", log_file)
    validator.validate_points_list()


if __name__ == '__main__':
    main()

