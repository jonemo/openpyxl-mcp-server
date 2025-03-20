import pytest
from pathlib import Path
from openpyxl_mcp_server import get_list_of_sheets


@pytest.mark.asyncio
async def test_get_list_of_sheets():
    # Get the path to the test file relative to this test file
    test_file = Path(__file__).parent / "testdata" / "simple_workbook.xlsx"
    result = await get_list_of_sheets(str(test_file))
    sheets = result.split("\n")
    assert len(sheets) == 2
    assert sheets[0] == "Name: First Worksheet, Dimensions: A1:F4"
    assert sheets[1] == "Name: 🧮, Dimensions: A1:B1"
