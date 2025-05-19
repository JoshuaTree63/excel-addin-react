import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class BalanceSheetTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Balance Sheet", base_location=base_location)

        self._member_name_value_map = dict(
            asset="Asset",
            cash_in_hand="Cash in hand",
            total_asset="Total",
            equity="Equity",
            retained_earnings="Retained Earnings",
            debt="Debt",
            total_equity_and_debt="Total",
        )
