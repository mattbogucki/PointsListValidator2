# PointsListValidator2

Points List Validator is a tool to validate SCADA points lists delivered by the EPC contractors against the Clearway Standards for all devices at a renewables power plant.

This code can be compiled into an executable using pyinstaller using the following command.

pyinstaller

The code can also be run directly with python3.  The Attachment3 document contains Clearway's standard naming conventions.  The code is designed so that even in executable form this document is linked at runtime so that changes can be made without the need to re-compile.  Therefore, when running the .exe Attachment 3 is expected to be in the same directory.

Syntax:

main.exe points_list.xlsx tab_of_excel_sheet

Example:

main.exe tsn1.xlsx 1

The code logs to stdout as well as a log file on the users desktop
