import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class PAndLTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="P&L", base_location=base_location)

        self._member_name_value_map = dict(
            gross_revenues="Gross revenues",
            opex="OPEX",
            ebitda="EBITDA",
            amortization="Amortization",
            ebit="EBIT",
            interest="Interest",
            ebt_taxable_profit="EBT (Taxable Profit)",
            income_tax="Income Tax",
            net_profit="Net Profit",
        )
