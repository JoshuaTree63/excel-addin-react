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
            traffic_passenger_car_pc="TRAFFIC - Passenger Car (PC)",
            traffic_heavy_vehicle_hv="TRAFFIC - Heavy Vehicle (HV)",
            revenue_passenger_car_pc="REVENUE - Passenger Car (PC)",
            revenue_heavy_vehicle_hv="REVENUE - Heavy Vehicle (HV)",
            total_revenue_real="Total Revenue (real)",
            total_revenue_nominal="Total Revenue (nominal)",
        ) 