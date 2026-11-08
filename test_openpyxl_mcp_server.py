import tempfile
from unittest.mock import patch

import pytest
from pathlib import Path
from openpyxl import Workbook
from openpyxl_mcp_server import get_cell_value
from openpyxl_mcp_server import get_list_of_sheets
from openpyxl_mcp_server import get_values_of_cell_range
from openpyxl_mcp_server import mcp
from openpyxl_mcp_server import resolve_path_and_assert_file_exists
from openpyxl_mcp_server import search_in_cell_range


@pytest.fixture
def workbook_with_coordinates():
    # Every cell in A1:AB12 contains its own coordinate as value, crossing both
    # the one-to-two digit row boundary and the one-to-two letter column boundary
    with tempfile.TemporaryDirectory() as temp_dir:
        wb = Workbook()
        sheet = wb.active
        sheet.title = "Sheet1"
        for row in sheet.iter_rows(min_row=1, max_row=12, min_col=1, max_col=28):
            for cell in row:
                cell.value = cell.coordinate
        test_file = Path(temp_dir) / "coordinates.xlsx"
        wb.save(test_file)
        yield str(test_file)


@pytest.mark.asyncio
async def test_get_list_of_sheets():
    # Get the path to the test file relative to this test file
    test_file = Path(__file__).parent / "testdata" / "simple_workbook.xlsx"
    result = await get_list_of_sheets(str(test_file))
    sheets = result.split("\n")
    assert len(sheets) == 2
    assert sheets[0] == "Name: First Worksheet, Dimensions: A1:F4"
    # Yes, you can have use emojis in sheet names:
    assert sheets[1] == "Name: 🧮, Dimensions: A1:B1"


@pytest.mark.asyncio
async def test_get_cell_value():
    test_file = Path(__file__).parent / "testdata" / "simple_workbook.xlsx"
    assert await get_cell_value(str(test_file), "First Worksheet", "B2") == "1"
    assert await get_cell_value(str(test_file), "First Worksheet", "C2") == "hello"


def test_resolve_path_and_assert_file_exists_full_path():
    test_file = Path(__file__).parent / "testdata" / "simple_workbook.xlsx"
    result = resolve_path_and_assert_file_exists(str(test_file))
    assert result == test_file


def test_resolve_path_and_assert_file_exists_in_home_folder():
    test_file_name = "file_in_home_folder.xlsx"
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        test_file_path = temp_dir_path / test_file_name
        test_file_path.touch()
        with patch("openpyxl_mcp_server.Path.expanduser", return_value=test_file_path):
            result = resolve_path_and_assert_file_exists(f"~/{test_file_name}")
            assert result == test_file_path


@pytest.mark.parametrize("home_folder_name", ["Desktop", "Downloads"])
def test_resolve_path_and_assert_file_exists_in_default_folder(home_folder_name):
    test_file_name = "file_in_desktop_folder.xlsx"
    # Make a temporary directory with the file, then mock Path.home() to return that directory
    with tempfile.TemporaryDirectory() as temp_dir:
        temp_dir_path = Path(temp_dir)
        home_folder_path = temp_dir_path / home_folder_name
        home_folder_path.mkdir(parents=True, exist_ok=True)
        test_file_path = home_folder_path / test_file_name
        test_file_path.touch()
        with patch("openpyxl_mcp_server.Path.home", return_value=temp_dir_path):
            result = resolve_path_and_assert_file_exists(test_file_name)
            assert result == test_file_path


@pytest.mark.asyncio
async def test_all_tools_have_descriptions():
    tools = await mcp.list_tools()
    assert len(tools) == 6
    for tool in tools:
        assert tool.description, f"Tool {tool.name} has no description"
        assert "The path to the Excel workbook." in tool.description


@pytest.mark.asyncio
@pytest.mark.parametrize(
    "top_left_cell, bottom_right_cell, expected_cells",
    [
        ("A1", "A3", ["A1", "A2", "A3"]),
        ("A9", "A11", ["A9", "A10", "A11"]),
        ("A11", "A9", ["A9", "A10", "A11"]),
        ("Z1", "AA1", ["Z1", "AA1"]),
        ("B1", "A2", ["A1", "B1", "A2", "B2"]),
        ("A1", "A1", ["A1"]),
    ],
)
async def test_get_values_of_cell_range(
    workbook_with_coordinates, top_left_cell, bottom_right_cell, expected_cells
):
    result = await get_values_of_cell_range(
        workbook_with_coordinates, "Sheet1", top_left_cell, bottom_right_cell
    )
    assert result == "\n".join(f"{cell}: {cell}" for cell in expected_cells)


@pytest.mark.asyncio
async def test_get_values_of_cell_range_invalid_cell(workbook_with_coordinates):
    with pytest.raises(ValueError):
        await get_values_of_cell_range(workbook_with_coordinates, "Sheet1", "A0", "A3")


@pytest.mark.asyncio
async def test_search_in_cell_range(workbook_with_coordinates):
    result = await search_in_cell_range(
        workbook_with_coordinates, "Sheet1", "A9", "A11", "A10"
    )
    assert result == "A10: A10"
