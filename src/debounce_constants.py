# Constants used to throttle expensive workbook‑shape scans
SHAPE_SCAN_EVERY_N_WRITES = 3
MAX_CONSECUTIVE_ERRORS = 2

STRUCTURAL_WRITE_TOOLS = {
    "create_sheet_tool",
    "delete_sheet_tool",
    "merge_cells_range_tool",
    "unmerge_cells_range_tool",
    "set_named_ranges_tool",
}

# Treat table insertion as a structural write so shape refresh logic
# fires deterministically after tables are added.
STRUCTURAL_WRITE_TOOLS |= {"insert_table_tool"}