import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class CostsTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Costs", base_location=base_location)

        self._member_name_value_map = dict(
            maintenance_and_spv_costs="Maintenance & SPV costs",
            total_costs_nominal="Total Costs (nominal)",
        ) 