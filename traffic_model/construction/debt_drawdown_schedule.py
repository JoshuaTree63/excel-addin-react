import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class DebtDrawdownScheduleTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Debt Drawdown Schedule", base_location=base_location)

        self._member_name_value_map = dict(
            debt_outstanding_bop_k_gbp="Debt Outstanding BoP (k£)",
            debt_drawdown_k_gbp="Debt Drawdown (k£)",
            debt_outstanding_eop_k_gbp="Debt Outstanding EoP (k£)",
            arrangement_fee_k_gbp="Arrangement fee (k£)",
            engagement_fee_k_gbp="Engagement fee (k£)",
            capitalized_interests_k_gbp="Capitalized Interests (k£)",
        ) 