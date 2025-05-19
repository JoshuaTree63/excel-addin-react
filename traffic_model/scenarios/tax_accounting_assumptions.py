import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class TaxAccountingAssumptionsTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Tax & Accounting Assumptions", base_location=base_location)

        self._member_name_value_map = dict(
            tax_rate="Tax Rate",
            dividend_distribution_policy="Dividend distribution policy",
        )
