from PointModels.Point import Point


class BinaryOutputPoint(Point):

    def __init__(self, row: int, substation: str, point_name: str, device_id: str, device_type: str, description: str,
                 dnp_index: str, availability: str, state_table: str, feedback_point_dnp_index: str):
        super().__init__(row, substation, point_name, device_id, device_type, description, dnp_index, availability)
        self._state_table = state_table
        self._feedback_point_dnp_index = feedback_point_dnp_index

    def get_state_table(self):
        return self._state_table

    def get_feedback_dnp_index(self):
        return self._feedback_point_dnp_index
