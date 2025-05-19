import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class DAndATable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="D&A", base_location=base_location)

        self._member_name_value_map = dict(
            asset_value_opening_k_gbp="Asset Value opening (k£)",
            asset_value_closing_k_gbp="Asset Value closing (k£)",
            amortization_k_gbp="Amortization (k£)",
        )
