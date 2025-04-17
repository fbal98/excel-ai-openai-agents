from agents import Agent, function_tool, RunContextWrapper, ModelSettings
from .tools import (
    get_sheet_names_tool,
    get_active_sheet_name_tool,
    set_cell_value_tool,
    get_cell_value_tool,
    get_range_values_tool,  # Tool for verifying ranges
    set_range_style_tool,
    create_sheet_tool,
    delete_sheet_tool,
    merge_cells_range_tool,
    unmerge_cells_range_tool,
    set_row_height_tool,
    set_column_width_tool,
    set_cell_formula_tool,
    set_cell_values_tool,  # Bulk tool
    set_table_tool,        # Bulk helper: write 2‑D table
    set_columns_widths_tool,  # Bulk helper: column widths
    set_range_formula_tool,   # Bulk helper: range formulas
    write_and_verify_range_tool,  # Composite write+verify
    get_cell_style_tool,          # Style inspectors
    get_range_style_tool,
    save_workbook_tool,
    CellValueMap,  # Import type for clarity if needed later
    CellStyle,     # Import type for clarity if needed later
)
from .context import AppContext

# Decorate tool functions with @function_tool and ensure detailed docstrings
get_sheet_names_tool = function_tool(get_sheet_names_tool, strict_mode=False)
get_active_sheet_name_tool = function_tool(get_active_sheet_name_tool, strict_mode=False)
set_cell_value_tool = function_tool(set_cell_value_tool, strict_mode=False)
get_cell_value_tool = function_tool(get_cell_value_tool, strict_mode=False)
get_range_values_tool = function_tool(get_range_values_tool, strict_mode=False)
set_range_style_tool = function_tool(set_range_style_tool, strict_mode=False)
create_sheet_tool = function_tool(create_sheet_tool, strict_mode=False)
delete_sheet_tool = function_tool(delete_sheet_tool, strict_mode=False)
merge_cells_range_tool = function_tool(merge_cells_range_tool, strict_mode=False)
unmerge_cells_range_tool = function_tool(unmerge_cells_range_tool, strict_mode=False)
set_row_height_tool = function_tool(set_row_height_tool, strict_mode=False)
set_column_width_tool = function_tool(set_column_width_tool, strict_mode=False)
set_cell_formula_tool = function_tool(set_cell_formula_tool, strict_mode=False)
set_cell_values_tool = function_tool(set_cell_values_tool, strict_mode=False)
set_table_tool = function_tool(set_table_tool, strict_mode=False)
set_columns_widths_tool = function_tool(set_columns_widths_tool, strict_mode=False)
set_range_formula_tool = function_tool(set_range_formula_tool, strict_mode=False)
save_workbook_tool = function_tool(save_workbook_tool, strict_mode=False)
write_and_verify_range_tool = function_tool(write_and_verify_range_tool, strict_mode=False)
get_cell_style_tool = function_tool(get_cell_style_tool, strict_mode=False)
get_range_style_tool = function_tool(get_range_style_tool, strict_mode=False)


SYSTEM_PROMPT="""
You are **Excel Assistant**, an autonomous spreadsheet expert.

GLOBAL MISSION
• Understand the user's natural‑language goal.  
• Manipulate the workbook solely through the function tools supplied in this environment.  
• Keep thinking & acting until the goal is fully achieved or you have exhausted safe options.  
• Produce a single, concise final answer or confirmation for the user—nothing else.

CORE LOOP  (executed internally; do NOT reveal)
1. **PLAN** – Think step‑by‑step: break the request into sub‑tasks and draft an ordered strategy.  
2. **ACT** – For each sub‑task, select the minimal tool(s) that move you closer to the goal and emit the corresponding function call(s).  
3. **OBSERVE** – Read the tool's output. Store observations in memory; avoid repeating the same read.  
4. **CRITIQUE & CORRECT** – Ask yourself:  
   • Did the observation match expectations?  
   • Is the overall goal now satisfied?  
   • If not, refine your plan, adjust parameters, or choose a different tool.  
   • On any error, parse the message, reason about likely cause, and retry up to **2** times with a corrected approach; if still blocked, ask the user for clarification.  
5. Repeat steps 1‑4 until the task is DONE or the framework halts you.

TERMINATION
• When confident the goal is met, output **DONE** (internally) and present the user with a short (≤ 2 sentences) result or confirmation—no tool traces, no step‑by‑step reasoning.

CRITICAL RULES
• **ONLY** use the listed tools; never invent new functions or fabricate spreadsheet data.  
• Double‑check any calculation or write‑operation by reading back or using a verification tool when available.  
• Keep internal thoughts private; expose only the final concise answer.  
• No apologies unless a hard limit prevents completion.  
• Maintain column/row/sheet integrity—never overwrite data unless the user asked for it or it is required by the goal.  
• Respect guardrails: if a tool or guardrail rejects an input, rethink before retrying.  
• When many *adjacent* cells/columns/rows are affected, always prefer bulk helpers
  (`set_table_tool`, `set_columns_widths_tool`, `set_range_formula_tool`) over
  loops of single‑cell calls. This keeps the job under the 25‑turn limit.
• If more information is needed (e.g., missing sheet name), ask a clarifying question instead of guessing.  

EXAMPLE (Q&A)
Q (user): "Fill A1:C3 with 1 2 3 / 4 5 6 / 7 8 9."
Thought (plan): Need bulk write.
Action: set_table_tool("Sheet1","A1",[[1,2,3],[4,5,6],[7,8,9]])
Observation: True
Answer: "Filled 9 cells."

STYLE FOR FINAL USER MESSAGE
• Direct, factual, and minimal:  
  – Example success: "The average of column B is **42.7**."  
  – Example confirmation: "Created 'Summary' sheet and filled the totals."  
• Mention next steps only if the user explicitly requested them.

(The function tool catalogue is documented separately; refer to it when planning.)

"""

# Configure model settings for improved performance
model_settings = ModelSettings(
    temperature=0.2,  # Lower temperature for more deterministic responses
    tool_choice="auto",  # Let the model decide when to use tools
)

excel_assistant_agent = Agent[AppContext]( # Specify context type for clarity
    name="Excel Assistant",
    instructions=SYSTEM_PROMPT,
    model="gpt-4.1-mini",  # Using the mini version for cost efficiency
    model_settings=model_settings,

    tools=[
        get_sheet_names_tool,
        get_active_sheet_name_tool,
        set_cell_value_tool,
        get_cell_value_tool,
        get_range_values_tool,    # Verify cell ranges
        set_range_style_tool,
        create_sheet_tool,
        delete_sheet_tool,
        merge_cells_range_tool,
        unmerge_cells_range_tool,
        set_row_height_tool,
        set_column_width_tool,
        set_cell_formula_tool,
        set_cell_values_tool,     # Bulk tool
        set_table_tool,           # Bulk helper: write 2‑D table
        set_columns_widths_tool,  # Bulk helper: column widths
        set_range_formula_tool,   # Bulk helper: range formulas
        write_and_verify_range_tool,  # Composite write + self‑check
        get_cell_style_tool,          # Style inspectors
        get_range_style_tool,
        save_workbook_tool,
    ],
)

# Example usage (for testing purposes, not part of the agent definition)
async def main():
    print("Excel AI Assistant agent is ready. (Run via CLI)")

# You can add a check block if needed for direct execution, but CLI is the main entry point
# if __name__ == "__main__":
#     import asyncio
#     asyncio.run(main())