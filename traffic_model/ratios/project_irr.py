import datetime
import typing

from sheets.scenarios.scenario_table import ScenarioTable

from ..base_table import BaseTable, TABLE_HEADER_STYLE
from ..item import TableItem, TableItemStyle
from .table_entry import ScenarioTableEntry


class ProjectIRRTable(ScenarioTable):
    def __init__(self, parent, base_location) -> None:
        super().__init__(parent=parent, name="Project IRR", base_location=base_location)

        self._member_name_value_map = dict(
            capex_incl_spv_costs="CAPEX (incl. SPV costs)",
            arrangement_fee="Arrangement fee",
            engagement_fee="Engagement fee",
            capitalized_interests="Capitalized Interests",
            total_investment="Total Investment",
            revenues_pc_hv="Revenues PC+HV",
            maintenance_and_spv_costs="Maintenance & SPV Costs",
            corporate_tax="Corporate Tax",
            cash_flow="Cash Flow",
            project_irr="Project IRR",
        ) 