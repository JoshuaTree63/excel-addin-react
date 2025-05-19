import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class TrafficRevenueAssumptionsTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Traffic & Revenue Assumptions", base_location=base_location)

        self._member_name_value_map = dict(
            traffic_passenger_car="Traffic - Passenger Car (PC)",
            traffic_heavy_vehicle="Traffic - Heavy Vehicule (HV)",
            traffic_evolution_per_year="Traffic Evolution per year (PC & HV)",
            inflation_per_year="Inflation per year",
            toll_rate_passenger_car="Toll Rate - Passenger Car (PC)",
            toll_rate_heavy_vehicle="Toll Rate - Heavy Vehicule (HV)",
        )
