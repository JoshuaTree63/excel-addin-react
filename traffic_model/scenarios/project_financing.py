import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class ProjectFinancingTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Project Financing", base_location=base_location)

        self._member_name_value_map = dict(
            repayment_start_date="Repayment Start Date",
            repayment_end_date="Repayment End Date",
            base_interest_rate="Base Interest Rate",
            fixed_rate_margin="Fixed Rate Margin",
            arrangement_fee="Arrangement fee",
            engagement_fee="Engagement fee",
            gearing="Gearing",
        )
