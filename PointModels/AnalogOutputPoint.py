from PointModels.Point import Point


class AnalogOutputPoint(Point):

    def __init__(self, row: int, substation: str, device_id: str, point_name: str, device_type: str, description: str,
                 dnp_index: str, availability: str, units: str, feedback_point: str):
        super().__init__(row, substation, point_name, device_id, device_type, description, dnp_index, availability)
        self._units = units
        self._feedback_point = feedback_point

    def get_units(self):
        return self._units

    def get_feedback_point(self):
        return self._feedback_point


