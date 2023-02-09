import openpyxl

from DeviceModels.Device import Device


class DeviceManager(object):

    def __init__(self, devices_excel_file):
        self._devices_excel_file = devices_excel_file
        self._device_dictionary = {}

    # Lazy Load as needed
    def create_device(self, device_type: str) -> Device:
        device_schema = self._device_dictionary.get(device_type, None)

        if device_schema:
            return Device(device_schema, device_type)
        else:
            wb = openpyxl.load_workbook(self._devices_excel_file, data_only=True)
            sheet = wb[device_type]
            schema = []
            for i, row in enumerate(sheet.values):
                if i == 0:
                    continue  # skip header
                if row[0]:
                    schema.append((row))

            self._device_dictionary[device_type] = schema
            return Device(schema, device_type)


