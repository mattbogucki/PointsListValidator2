class Point(object):

    def __init__(self, row: int, substation: str, point_name: str, device_id: str,  device_type: str, description: str,
                 dnp_index: str, availability: str):
        self._row = row
        self._substation = substation
        self._device_id = device_id
        self._point_name = point_name
        self._device_type = device_type
        self._description = description
        self._dnp_index = dnp_index
        self._availability = availability

    def get_row(self) -> int:
        return self._row

    def get_substation(self) -> str:
        return self._substation

    def get_point_name(self) -> str:
        return self._point_name

    def get_device_type(self) -> str:
        return self._device_type

    def get_source_device(self) -> str:
        return self._device_id

    def get_description(self) -> str:
        return self._description

    def get_dnp_index(self) -> str:
        return self._dnp_index

    def get_availability(self) -> str:
        return self._availability
