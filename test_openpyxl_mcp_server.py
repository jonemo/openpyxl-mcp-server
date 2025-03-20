import tempfile
from unittest.mock import patch

import pytest
from pathlib import Path
from openpyxl_mcp_server import get_list_of_sheets
from openpyxl_mcp_server import resolve_path_and_assert_file_exists


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
