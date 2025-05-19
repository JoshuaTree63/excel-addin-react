import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class DebteTimelineTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Debte Timeline", base_location=base_location)

        self._member_name_value_map = dict(
            availability_start_date="Availability Start Date",
            availability_active="Active",
            availability_end_date="Availability End Date",
            repayment_start_date="Repayment Start Date",
            repayment_active="Active",
            repayment_end_date="Repayment End Date",
        ) 