import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class CashflowStatementTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Cashflow Statement", base_location=base_location)

        self._member_name_value_map = dict(
            gross_revenues="Gross Revenues",
            opex_maintenance_and_spv_costs="OPEX (Maintenance & SPV Costs)",
            income_tax="Income Tax",
            cfads="Cash Flow Available for Debt Service (CFADS)",
            interest="Interest",
            principal_repayment="Principal repayment",
            cash_flow_available_for_equity="Cash Flow Available for Equity",
            dividends="Dividends",
            net_cash_flow="Net Cash Flow",
            cash_in_hands="Cash in hands",
        )
