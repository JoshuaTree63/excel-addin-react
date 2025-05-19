import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class UsesOfFundsForCircularityBreakdownPurposeTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Uses of Funds for Circularity Breakdown Purpose", base_location=base_location)

        self._member_name_value_map = dict(
            uses_total_k_gbp="Uses Total (k£)",
            total_debt_amount_k_gbp="Total Debt Amount (k£)",
            total_equity_amount_k_gbp="Total Equity Amount (k£)",
        ) 