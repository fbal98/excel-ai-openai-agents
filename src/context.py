from dataclasses import dataclass, field
from .excel_ops import ExcelManager

@dataclass
class AppContext:
    excel_manager: ExcelManager
    # Generic bag for planner / retry metadata
    state: dict = field(default_factory=dict)