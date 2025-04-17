from dataclasses import dataclass
from .excel_ops import ExcelManager

@dataclass
class AppContext:
    excel_manager: ExcelManager
    # Add any other shared state/dependencies here if needed later
