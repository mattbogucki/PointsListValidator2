from PointModels.Point import Point


class AnalogInputPoint(Point):

    def __init__(self, row: int, substation: str, device_id: str, point_name: str, device_type: str, description: str,
                 dnp_index: str, availability: str, units: str, egu_min: str, egu_max: str):

        super().__init__(row, substation, point_name, device_id, device_type, description, dnp_index, availability)
        self._units = units
        self._egu_min = egu_min
        self._egu_max = egu_max

    def get_units(self):
        return self._units

    def get_egu_min(self):
        return self._egu_min

    def get_egu_max(self):
        return self._egu_max
