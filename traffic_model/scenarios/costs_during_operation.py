import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class CostsDuringOperationTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Costs During Operation", base_location=base_location)

        self._member_name_value_map = dict(
            maintenance_including_heavy_maintenance_and_spv_costs="Maintenance (including heavy maintenance & SPV costs)",
            inflation_per_year_costs="Inflation per year (costs) from beginning of concession",
        ) 