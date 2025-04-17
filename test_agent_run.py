import asyncio
from agents import Runner
from src.agent_core import excel_assistant_agent
from src.excel_ops import ExcelManager
from src.context import AppContext

async def run_test(instruction: str, input_file: str = None, output_file: str = "test_output.xlsx"):
    excel_manager = ExcelManager(file_path=input_file) if input_file else ExcelManager()
    app_context = AppContext(excel_manager=excel_manager)
    print(f"Running test with instruction: {instruction}")
    result = await Runner.run(excel_assistant_agent, input=instruction, context=app_context, max_turns=25)
    print(f"Final output: {result.final_output}")
    saved = excel_manager.save_workbook(output_file)
    print(f"Workbook saved: {saved} to {output_file}")

async def main():
    # Simple table creation with default columns
    await run_test("Create a table of 5 users with default columns", output_file="output1.xlsx")

    # Table creation with specified columns
    await run_test("Create a table of 3 persons with columns Name, Age, Email", output_file="output2.xlsx")

    # Ambiguous sheet name, expect minimal clarifying question
    await run_test("Add a summary sheet with totals", output_file="output3.xlsx")

    # Complex multi-step instruction
    await run_test("Create a report sheet, add a table of 10 sales records, apply bold style to header row", output_file="output4.xlsx")

if __name__ == "__main__":
    asyncio.run(main())