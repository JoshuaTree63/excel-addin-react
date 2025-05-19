import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class SourcesTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Sources", base_location=base_location)

        self._member_name_value_map = dict(
            equity_k_gbp="Equity (k£)",
            debt_k_gbp="Debt (k£)",
            total_sources_k_gbp="Total Sources (k£)",
        ) 