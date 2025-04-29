import logging
from typing import Optional, Tuple, Any # Added Tuple, Any

from agents import Agent, function_tool, FunctionTool, RunContextWrapper, Runner, Usage, RunResult # Added Runner, Usage, RunResult
# Import all tools from the new tools package via its __init__.py
# This imports all functions/classes listed in src.tools.__all__
from . import tools as excel_tools

# Import specific types needed directly if not re-exported or for clarity
from .tools.core_defs import CellValueMap, CellStyle

# Import other necessary components
from .context import AppContext, WorkbookShape # Ensure WorkbookShape is imported if used directly
from .hooks import SummaryHooks # Changed ActionLoggingHooks to SummaryHooks
from .costs import dollars_for_usage # Added import
from .conversation_context import ConversationContext # Added import
# tool_wrapper logic is handled implicitly by the FunctionTool decorator or manual application in tools/__init__.py

logger = logging.getLogger(__name__) # Added logger

# --- Configuration ---
MAX_SHEETS_IN_PROMPT = 30
MAX_HEADERS_PER_SHEET = 50  # Limit number of headers per sheet in the prompt

def _compact_headers(headers):
    """
    Compacts header representations to reduce token usage.
    Converts repetitive empty headers to a more compact form.

    Examples:
    - ["Name", "", "", "", "Date"] -> ["Name", "...", "Date"]
    - ["", "", "", ""] -> [] (all empty case)
    - ["Name", "Age", "Email"] -> ["Name", "Age", "Email"] (no change needed)
    """
    if not headers:
        return []

    # If we already have few headers or none are empty, return as is
    empty_count = sum(1 for h in headers if not h)
    if len(headers) <= 10 or empty_count == 0:
        return headers

    # If all headers are empty, return an empty list
    if empty_count == len(headers):
        return []

    # Compact format: find meaningful headers and compress empty spans
    result = []
    empty_streak = 0

    for i, header in enumerate(headers):
        if not header:  # Empty header
            empty_streak += 1
            # Only add ellipsis if we've seen 3+ consecutive empty headers
            # and haven't already added an ellipsis
            if empty_streak == 3 and (not result or result[-1] != "..."):
                result.append("...")
        else:  # Non-empty header
            empty_streak = 0
            result.append(str(header)) # Ensure string conversion

    return result

def _format_workbook_shape(shape: Optional[WorkbookShape]) -> str: # Use WorkbookShape directly
    """Formats the WorkbookShape into a string for the prompt, respecting limits."""
    import logging
    logger = logging.getLogger(__name__)

    if not shape:
        # Treat the first, shape‑less scan as version 1 so later math never sees v=0.
        return "<workbook_shape v=1></workbook_shape>"

    # Limit sheets included in the prompt
    limited_sheets = list(shape.sheets.items())[:MAX_SHEETS_IN_PROMPT]
    limited_sheet_names = {s_name for s_name, _ in limited_sheets} # Set for faster lookup
    limited_headers = {s: h for s, h in shape.headers.items() if s in limited_sheet_names}
    # Named ranges are usually fewer, include all for now
    named_ranges = shape.names.items()

    sheets_str = '; '.join(f'{s}:{rng}' for s, rng in limited_sheets) if limited_sheets else ""

    # Process headers - compact them and limit per sheet
    processed_headers = {}
    total_original_headers = 0
    total_compacted_headers = 0

    for sheet_name, headers in limited_headers.items():
        original_len = len(headers)
        total_original_headers += original_len

        # First compact the headers to reduce empty spans
        compacted = _compact_headers(headers)
        compacted_len_after_compact = len(compacted)

        # Then limit to maximum number of headers if still large
        if compacted_len_after_compact > MAX_HEADERS_PER_SHEET:
            # Keep first and last few headers with an ellipsis in between
            front_count = MAX_HEADERS_PER_SHEET // 2
            back_count = MAX_HEADERS_PER_SHEET - front_count # Adjust back count
            front_headers = compacted[:front_count]
            back_headers = compacted[-back_count:]
            compacted = front_headers + ["..."] + back_headers
            logger.debug(f"Sheet '{sheet_name}': Headers truncated from {compacted_len_after_compact} to {len(compacted)} due to MAX_HEADERS_PER_SHEET limit")

        processed_headers[sheet_name] = compacted
        final_compacted_len = len(compacted)
        total_compacted_headers += final_compacted_len

        # Log individual sheet stats only if significant changes occurred
        if original_len > 10 and original_len != final_compacted_len:
             logger.debug(f"Sheet '{sheet_name}': Headers optimized from {original_len} to {final_compacted_len} items")

    # Log overall stats
    if total_original_headers > 0:
        reduction_percent = 0
        if total_original_headers > 0: # Avoid division by zero
             reduction_percent = ((total_original_headers - total_compacted_headers) / total_original_headers) * 100
        logger.info(f"Workbook shape optimization: Reduced headers from {total_original_headers} to {total_compacted_headers} items ({reduction_percent:.1f}% reduction)")

    headers_str = '; '.join(f'{s}:{",".join(map(str,h))}' for s, h in processed_headers.items() if h) if processed_headers else ""

    # Include named ranges
    names_str = '; '.join(f'name:{n}={ref}' for n, ref in named_ranges) if named_ranges else ""

    # Assemble the final string, filtering empty sections
    shape_content_parts = [part for part in [sheets_str, headers_str, names_str] if part]
    shape_content = '\n'.join(shape_content_parts)

    final_shape = f"<workbook_shape v={shape.version}>\n{shape_content}\n</workbook_shape>"
    logger.debug(f"Final workbook shape size: {len(final_shape)} characters")

    return final_shape


def _dynamic_instructions(wrapper: RunContextWrapper[AppContext], agent: Agent) -> str:  # noqa: D401
    """
    Return the static SYSTEM_PROMPT.
    Dynamic context (shape, progress, state) is now injected directly into
    the conversation history by hooks and ConversationContext.
    """
    # No longer dynamically adding shape, summary, or session state here.
    # This info is now managed within the conversation_history.
    return SYSTEM_PROMPT


SYSTEM_PROMPT="""
You are a highly skilled professional with exceptional organizational abilities and meticulous attention to detail. Your expertise encompasses project management, data analysis, and advanced Excel techniques, enabling you to create sophisticated spreadsheets, streamline workflows, and generate comprehensive reports effortlessly. You excel at optimizing processes, identifying trends, and solving complex problems in various business contexts.

As an interactive assistant, you work with users in creating and managing Excel sheets and project management tasks. Leveraging your extensive knowledge and available tools, you provide tailored solutions that are both functionally robust and visually polished. Your approach combines the most efficient formulas, pivot tables, and data visualization techniques to deliver professional, aesthetically pleasing spreadsheets that enhance productivity and decision-making.

# Tone and style
You should be concise, direct, and to the point. When you are about to run a non-trivial data manipulation or analysis, you should explain what the operation does and why you are running it to the user before calling the tool, to make sure the user understands what you are doing.
Remember that your output will be displayed on a command line interface. Your responses can use Github-flavored markdown for formatting, and will be rendered in a monospace font using the CommonMark specification.
Output text to communicate with the user; all text you output outside of tool use is displayed to the user. Only use tools to complete tasks. Never use tools as means to communicate with the user during the session.
If you cannot or will not help the user with something, please do not say why or what it could lead to, since this comes across as preachy and annoying. Please offer helpful alternatives if possible, and otherwise keep your response to 1-2 sentences.
IMPORTANT: You should minimize output tokens as much as possible while maintaining helpfulness, quality, and accuracy. Only address the specific query or task at hand, avoiding tangential information unless absolutely critical for completing the request. If you can answer in 1-3 sentences or a short paragraph, please do.
IMPORTANT: You should NOT answer with unnecessary preamble or postamble (such as explaining your code or summarizing your action), unless the user asks you to.
IMPORTANT: Keep your responses short, since they will be displayed on a command line interface. You MUST answer concisely with fewer than 4 lines (not including tool use or data generation), unless user asks for detail. Answer the user's question directly, without elaboration, explanation, or details. One word answers are best. Avoid introductions, conclusions, and explanations. You MUST avoid text before/after your response, such as "The answer is <answer>.", "Here is the content of the file..." or "Based on the information provided, the answer is..." or "Here is what I will do next...". Here are some examples to demonstrate appropriate verbosity:
IMPORTANT: You MUST always explain in a sentence why you are calling a tool before calling it.

<example>
user: create a new sheet called "movies" 
assistant: I will create a new sheet called "movies".
assistant: [calls create_sheet_tool]
assistant: New sheet is created. 
</example>

<example>
user: Apply bold style to the header row
assistant: I will apply bold style to the header row.
assistant: [calls apply_bold_to_row]
assistant: Bold style has been applied to the header row.
</example>

<example>
user: Add a SUM formula to column C
assistant: I will add a SUM formula to column C.
assistant: [calls add_sum_formula]
assistant: SUM formula has been added to column C.
</example>

<example>
user: Copy data from Sheet1 to Sheet2
assistant: I will copy data from Sheet1 to Sheet2.
assistant: [calls copy_sheet_data]
assistant: Data has been copied from Sheet1 to Sheet2.
</example>

# Proactiveness
You are allowed to be proactive, but only when the user asks you to do something. You should strive to strike a balance between:
1. Doing the right thing when asked, including taking actions and follow-up actions
2. Not surprising the user with actions you take without asking
For example, if the user asks you how to approach something, you should do your best to answer their question first, and not immediately jump into taking actions.

# Sheets styling
- If you feel the user has a single table, or basic data visualization using cells only, you should apply a basic style to it give it a professional look.
- Do not style unclear or uncomplete user requests/sessions.

# Doing tasks
The user will primarily request you perform excel and sheets tasks. This includes working on empty sheets (from scratch), adding new data, revamping existing sheets, explaining data and doing calculations, and more. For these tasks the following steps are recommended:
1. Use the available scanning tools to understand the sheets and the user's query. You are encouraged to use the scanning and awareness tools extensively both in parallel and sequentially.
2. Implement the solution using all tools available to you. even if it feels it is not the best solution, do it. IMPORTANT: You must not tell the user if a solution is not the best one just do whatever you are asked with the tools.
3. Verify the solution if possible with tools if you can. NEVER assume solution is correct or it was applied. Always trust tool results.
4. VERY IMPORTANT: You must scan the sheet to verify the solution.
5. when you are moving big parts of the sheet, cut it and paste it in an empty space to persist it. 
6. Always favor doing data in bulk instead of row by row. 
7. very important: if you foresee a huge amount of data, to be processesed at once. don't bulk it in one call but call the buld tools multiple times. 
<example>
user: fill the sheet with random 10x1000 data   
assistant: I will fill the sheet with data of 10x1000.
assistant: [thinks internally to generate data]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A1", rows=rows[0:100])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A101", rows=rows[100:200])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A201", rows=rows[200:300])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A301", rows=rows[300:400])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A401", rows=rows[400:500])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A501", rows=rows[500:600])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A601", rows=rows[600:700])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A701", rows=rows[700:800])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A801", rows=rows[800:900])]
assistant: [calls set_table_tool(sheet="Sheet1", top_left="A901", rows=rows[900:1000])]
assistant: [calls get_range_values_tool(sheet="Sheet1", range_address="A1:J1000")]
assistant: I have filled the sheet with data of 10x1000.
</example>

# Tool usage policy
- When you are fulfilling a user request, always use the most context-efficient tool to minimize context usage and maximize relevant results.
- When executing multiple independent tool calls (such as reading, writing, or editing different sheet regions), always batch them for parallel execution. For example, if you need to update several cell ranges or perform multiple verification steps, run these in a single parallel batch for speed and efficiency.
- Always favor bulk operations (e.g., set_table_tool, set_cell_values_tool) over iterative or row-by-row updates to improve efficiency and reliability.
- After making changes, use the appropriate verification tool (e.g., get_range_values_tool) to confirm results; never assume success without checking tool output.
- You MUST answer concisely with fewer than 4 lines of text (excluding tool use or code generation), unless the user requests more detail.
"""

# Define the agent - Ensure all tools listed here are decorated in tools.py


# --- Define the list of tools for the agent ---
# Tools are imported via `excel_tools` alias from src/tools/__init__.py
# Ensure all tools intended for the agent are listed in src/tools/__all__
# and are decorated with @function_tool in their respective files.
_agent_tools_list = [
    excel_tools.open_workbook_tool,
    excel_tools.snapshot_tool,
    excel_tools.revert_snapshot_tool,
    excel_tools.get_sheet_names_tool,
    excel_tools.get_active_sheet_name_tool,
    excel_tools.set_cell_value_tool,
    excel_tools.get_cell_value_tool,
    excel_tools.get_range_values_tool,
    excel_tools.find_row_by_value_tool,
    excel_tools.get_dataframe_tool,
    excel_tools.set_range_style_tool,
    excel_tools.set_cell_style_tool,
    excel_tools.create_sheet_tool,
    excel_tools.delete_sheet_tool,
    excel_tools.merge_cells_range_tool,
    excel_tools.unmerge_cells_range_tool,
    excel_tools.set_row_height_tool,
    excel_tools.set_column_width_tool,
    excel_tools.set_columns_widths_tool,
    excel_tools.set_range_formula_tool,
    excel_tools.set_cell_formula_tool,
    excel_tools.set_cell_values_tool,
    excel_tools.set_table_tool, # Simple data write
    excel_tools.insert_table_tool, # Formatted Excel Table object
    excel_tools.set_rows_tool,
    excel_tools.set_columns_tool,
    excel_tools.set_named_ranges_tool,
    excel_tools.append_table_rows_tool,
    excel_tools.copy_paste_range_tool,
    excel_tools.write_and_verify_range_tool,
    excel_tools.get_cell_style_tool,
    excel_tools.get_range_style_tool,
    excel_tools.save_workbook_tool,
]

# The @function_tool decorator handles wrapping, so explicit conversion might not be needed
# if all tools are correctly decorated. If any raw functions were missed, this ensures they are converted.
_validated_agent_tools: list[FunctionTool] = [
    t if isinstance(t, FunctionTool) else function_tool(t) for t in _agent_tools_list
]

# --- Static agent definition removed ---
# excel_assistant_agent = Agent[AppContext](...) <-- REMOVED

# --- Updated run_and_cost ---
from .costs import dollars_for_usage
from agents import Runner, Agent, Usage
from typing import Tuple, Any
import logging # Added for logging

logger = logging.getLogger(__name__) # Added logger

async def run_and_cost(
    agent: Agent[AppContext], # The agent instance created by create_excel_assistant_agent
    *,
    input: str,
    context: AppContext, 
    **kw,
) -> Tuple[Any, Usage, float]:
    """
    Convenience helper: Runs the agent and calculates cost using litellm.
    Returns (result, usage, dollars).
    Stores cost/usage info into context.state if possible.
    """
    res = await Runner.run(agent, input=input, context=context, **kw, max_turns=25)

    # Assuming context object collects/updates usage; If not, need to adjust.
    # Check if usage is directly on context or nested
    usage = None
    if hasattr(context, 'usage') and isinstance(context.usage, Usage):
         usage = context.usage
    elif hasattr(context, 'state') and isinstance(context.state, dict) and 'usage' in context.state and isinstance(context.state['usage'], Usage):
         usage = context.state['usage'] # Example if usage is stored in state

    if not usage:
        logger.warning("Could not find Usage object in context after agent run. Cost calculation skipped.")
        cost = 0.0
        usage = Usage() # Create empty usage to avoid downstream errors
    else:
        # Get the actual model name string used by the agent instance for costing
        # Now agent.model should be the string itself (e.g., "gpt-4.1-mini" or "litellm/gemini/...")
        model_name_used = None
        if hasattr(agent, 'model') and isinstance(agent.model, str):
            model_name_used = agent.model
            logger.info(f"Found model name in agent.model: '{model_name_used}'")
        else:
            # Enhanced logging to better understand the structure
            logger.warning(f"Agent model attribute is not a string: {type(agent.model)}. Value: {agent.model}")
            
            # Try harder to get the model name - inspect agent attributes
            for attr_name in dir(agent):
                if attr_name.startswith('_'):
                    continue
                try:
                    attr_value = getattr(agent, attr_name)
                    logger.debug(f"Agent attr '{attr_name}': {type(attr_value)}")
                    if isinstance(attr_value, str) and "gemini" in attr_value.lower():
                        model_name_used = attr_value
                        logger.info(f"Found potential model name in agent.{attr_name}: '{model_name_used}'")
                        break
                except Exception:
                    pass


        if not model_name_used:
             logger.warning("Could not determine model name string from agent instance for cost calculation in run_and_cost.")
             cost = 0.0
        else:
            # Pass the specific model name string used in the run to the updated costing function
            # dollars_for_usage should now handle litellm/ prefixes if necessary
            cost = dollars_for_usage(usage, model_name_from_agent=model_name_used)

    # Store cost/usage details in context.state for CLI to potentially access
    if hasattr(context, 'state') and isinstance(context.state, dict):
        total_tokens = (getattr(usage, "input_tokens", 0) or 0) + (getattr(usage, "output_tokens", 0) or 0)
        context.state['last_run_cost'] = cost
        context.state['last_run_usage'] = {
            'input_tokens': getattr(usage, "input_tokens", 0) or 0,
            'output_tokens': getattr(usage, "output_tokens", 0) or 0,
            'total_tokens': total_tokens,
            'model_name': model_name_used or "Unknown" # Store the model name used if found
        }
        logger.debug(f"Stored run cost (${cost:.6f}) and usage ({total_tokens} tokens) for model '{model_name_used}' in context state.")
    else:
        logger.warning("Cannot store cost/usage in context.state (context missing 'state' dict).")

    # Append the raw result messages (user input + assistant output) to the history
    from agents.results import RunResult # Moved import here
    from .conversation_context import ConversationContext # Import the new helper

    if isinstance(res, RunResult):
        ctx_msgs: list = context.state.setdefault("conversation_history", [])
        # Use res.new_items which contains the actual generated items (tool calls, outputs, messages)
        # instead of res.to_input_list() which might include the initial input again.
        # We need to filter/format these items appropriately for history.
        # Let's keep it simple for now and use to_input_list, but filter duplicates later if needed.
        # Note: to_input_list() returns the original input PLUS new items.
        # This might lead to duplicates if history already contains the input.
        # A better approach might be to only add `res.new_items`.

        # Let's try adding only new_items, assuming the input is already in history from the previous turn or CLI handling.
        # We need to convert RunItem objects to the dict format expected by history.
        new_history_items = []
        for item in res.new_items:
            # Simple conversion, might need refinement based on item types
            if hasattr(item, 'to_dict'): # Check if item has a dict representation
                 item_dict = item.to_dict()
                 # Ensure 'role' and 'content' exist, adjust as needed based on item types
                 role = item_dict.get('role', 'assistant') # Default role
                 content = item_dict.get('content', str(item_dict)) # Default content
                 # TODO: Properly extract role/content based on item.type (MessageOutputItem, ToolCallItem etc.)
                 if role and content: # Basic check
                     new_history_items.append({'role': role, 'content': content})
            else:
                # Fallback for items without to_dict
                new_history_items.append({'role': 'system', 'content': f'<item type={item.type}>{str(item)}</item>'})

        # Avoid adding duplicates if the last message is identical
        if ctx_msgs and new_history_items and ctx_msgs[-1] == new_history_items[0]:
             new_history_items = new_history_items[1:] # Skip duplicate first item

        ctx_msgs.extend(new_history_items)
        logger.debug(f"Extended conversation_history with {len(new_history_items)} new items from RunResult.")

        # Now prune the potentially extended history
        ConversationContext.maybe_prune(context) # Pass the AppContext instance
    else:
        logger.warning("Result object is not a RunResultBase instance, cannot update conversation history.")


    return res, usage, cost

# --- Removed main execution block ---
# if __name__ == "__main__":
#       pass