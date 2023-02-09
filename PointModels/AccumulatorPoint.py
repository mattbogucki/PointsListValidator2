from PointModels.Point import Point


class AccumulatorPoint(Point):

    def __init__(self, row: int, substation: str, device_id: str, point_name: str, device_type: str, description: str,
                 dnp_index: str, availability: str, units: str, rollover: str):
        super().__init__(row, substation, point_name, device_id, device_type, description, dnp_index, availability)
        self._units = units
        self._rollover = rollover

    def get_units(self):
        return self._units

    def get_rollover_value(self):
        return self._rollover


