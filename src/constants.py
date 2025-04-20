# src/constants.py
# Shared constants to avoid circular imports between context.py and hooks.py

WRITE_TOOLS = {
    "open_workbook_tool",
    "set_cell_value_tool",
    "set_range_style_tool",
    "set_cell_style_tool",
    "create_sheet_tool",
    "delete_sheet_tool",
    "merge_cells_range_tool",
    "unmerge_cells_range_tool",
    "set_row_height_tool",
    "set_column_width_tool",
    "set_columns_widths_tool",
    "set_cell_formula_tool",
    "set_cell_values_tool",
    "set_table_tool",
    "insert_table_tool",
    "set_rows_tool",
    "set_columns_tool",
    "set_named_ranges_tool",
    "copy_paste_range_tool",
    "write_and_verify_range_tool",
    "revert_snapshot_tool",
    "append_table_rows_tool",
    # Exclude save_workbook_tool as it doesn't change the structure/content itself
}
