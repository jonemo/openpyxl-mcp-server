# openpyxl_mcp_server

## What is this?

A thin wrapper around the OpenPyXl Python library that exposes some of its features as Model Context Protocol (MCP) server. This allows Claude and other MCP clients to fetch data from Excel files.

Related links:

- https://openpyxl.readthedocs.io/en/stable/index.html
- https://modelcontextprotocol.io/quickstart/server

## Setting it up

Clone repo into some folder, in my examples below I am using `/Users/jonas/repos/openpyxl_mcp_server`. Then:

```sh
python -m venv venv
source venv/bin/activate
pip install -r requirements.txt
```

## Configuring Claude Desktop

See the "Testing your server with Claude for Desktop" section of https://modelcontextprotocol.io/quickstart/server for instructions how to configure Claude Desktop. The exact command to use for this server is below.

## Claude Desktop configuration

Note that this only works with the [Claude Desktop app](https://claude.ai/download), not with Claude running in the web browser.

If you have not used MCP with Claude Desktop before, you must create the `claude_desktop_config.json` configuration file.
You can do this by opening the "Settings" menu and then go to the "Developer" menu item, then follow the instructions.

Alternatively, you can create the file yourself. The common locations of it are:

- macOS: ~/Library/Application Support/Claude/claude_desktop_config.json
- Windows: %APPDATA%\Claude\claude_desktop_config.json

Add the following to your `claude_desktop_config.json`.
If you already use other MCP servers, you'll have to merge the JSON into the existing contents.

```json
{
  "mcpServers": {
    "inspect_excel_xlsx_file": {
      "command": "/Users/jonas/repos/openpyxl_mcp_server/venv/bin/python",
      "args": ["/Users/jonas/repos/openpyxl_mcp_server/openpyxl_mcp_server.py"]
    }
  }
}
```

After updating your configuration file, you need to restart the Claude app for the changes to take effect.

Note that there is a separate `config.json` file that lives in the same folder. Don't touch that one.

If you encounter problems, please consult [Anthropic's documentation](https://modelcontextprotocol.info/docs/quickstart/user/).
Someone also made a config file generator at https://claudedesktopconfiggenerator.com/

## Configuring Cursor to use this

Instructions for setting it up in Cursor are at https://docs.cursor.com/context/model-context-protocol.

## Example prompts

### Example 1

Uses the file `KimbroughEvelyn_A-5qg0_Dataset_20160624.xlsx` from the "Near-Port Air Quality Assessment Utilizing a Mobile Monitoring Approach" dataset from https://catalog.data.gov/dataset/near-port-air-quality-assessment-utilizing-a-mobile-monitoring-approach

From a quick glance at description and file contents it wasn't obvious to me what ports this dataset is about. Let's ask Claude:

> I have an Excel file in ~/Downloads/KimbroughEvelyn_A-5qg0_Dataset_20160624.xlsx that contains data associated with some study about air pollution near one or more ports in the United States. Can you infer from the file contents which ports are covered by this study?

https://claude.ai/share/d87f67a9-dbc0-4223-86ff-3b4c1f9fa2e7

After 11 calls to the MCP server, the response is:

> Looking at the latitude and longitude coordinates from the Fig 2 sheet and combining this with the location information, I can infer that this study is focused on the Port of Charleston, South Carolina.
>
> The coordinates in the dataset (approximately latitude 32.83 and longitude -79.87) correspond to the Charleston area, and the references to downtown, rail yard, and airport monitoring stations are consistent with air quality monitoring around a port city.
>
> This dataset appears to contain air pollution measurements (PM2.5 and PM10) collected near the Port of Charleston, with comparisons between port locations and background/reference sites.

### Example 2:

I'm the treasurer of a small 10-unit home owners association. Our book keeping happens in an Excel file in Google Sheets. The file has 9 sheets with nine sheets for things like the transactions list, members' dues accounts, quarterly reports.

The file contains non-public information, therefore this example only shows the prompt, a few key "insights", and the final answer:

> ~/Downloads/CASH FLOW 2024.xlsx are the finances of my home owner association. How much did we get billed for sprinkler system maintenance in 2024? How many transactions contributed to this total? When did each transaction happen?
>
> [...]
>
> I see a "Transactions" sheet that likely contains the detailed expense records. Let me search for "sprinkler" in this sheet to find relevant transactions:
>
> [...]
>
> Let me examine the YEAR 2024 summary sheet to see if sprinkler maintenance is listed as a category:
>
> [...]
>
> Great! I found "Sprinkler Maintenance" listed as a category in row 24 of the YEAR 2024 sheet. Now, let me check the actual transactions related to sprinkler maintenance:
>
> [...]
>
> Let me check more entries by searching specifically for "Sprinkler Maintenance" in the entire Transactions sheet:
>
> [...]
>
> Based on my analysis of your HOA's financial records in the "CASH FLOW 2024.xlsx" file, here's what I found about the sprinkler system maintenance expenses in 2024:
>
> Total amount billed for sprinkler system maintenance in 2024: $565.00
> Number of transactions: 1 transaction
> Transaction date: September 9, 2024
>
> The expense appears as a single transaction in the Transactions sheet. The spreadsheet also contains a budgeted amount for sprinkler maintenance of $750.00 for the year > in the "YEAR 2024 Actuals vs Planned" sheet, with a note stating "Not received any invoices yet" (which may be outdated since there is now one invoice recorded).
> Would you like me to analyze any other aspects of your HOA's sprinkler maintenance expenses or other financial information from the spreadsheet?

True story. Our sprinkler guy only billed us for the one visit in the spring and forgot to send the bill for the work later in the year. Don't worry, he got paid but it's booked in 2025.

### Example 3: Filling out a form

My Google searches for interesting Excel workbooks turned up form DPR 973 by the California Natural Resources Agency:
https://www.parks.ca.gov/pages/22374/files/Annual%20Report%20DPR973.xls This appears to be a form that gets filled out
by "associations" that support the California State Parks. To be clear, I am not affiliated with such an association.
But I converted the form from .xls format to .xlsx and asked Claude to help me fill it out:

> I need to file DPR973 because I work for a "cooperating association" for California State Parks. The form I have to fill out is this Excel sheet: /Users/jonas/Downloads/Annual Report DPR973.xlsx. We supported the State Parks by publishing books and generated revenue by selling firewood. Where in the workbook do I have to fill out cells?

After 11 calls to the tool the response seems correct, although I don't understand why Claude did not look at the remaining sheets at all:

> Based on my examination of the DPR973 form, I'll now provide guidance on where you need to fill out cells for your cooperating association that publishes books and sells firewood:
>
> 1. **Page 1**:
>
>    - You'll need to fill out the Association Information (Part I) with your organization's details (name, address, contact information)
>    - Update the year in cell C4 to the current reporting year
>
> 2. **Page 2 (A)** - This is where you'll record your sales information:
>
>    - For your book sales (which would be considered interpretive/educational):
>
>      - Line 1 (Cell F22): Gross Sales - Materials (for your published books)
>      - Line 2 (Cell F23): Any returns or refunds
>      - Line 3 (Cell F24): Cost of Goods Sold for the books
>
>    - For your firewood sales:
>
>      - Line 23 (Cell F51): Gross Sales - Firewood
>      - Line 24 (Cell F52): Any refunds/returned/damaged firewood
>      - Line 25 (Cell F53): Cost of Goods Sold for firewood
>
>    - The form will automatically calculate totals in cells F25 (book profits), F54 (firewood profits), and F58 (total net profits)
>
> 3. **Page 3** - Financial Summary:
>
>    - Your sales figures will automatically populate from Page 2(A) into:
>
>      - Line 6 (Cell F11): Total gross sales
>      - Line 7 (Cell F12): Total cost of goods/services + returns/refunds
>      - Line 8 (Cell H13): Net profits for all materials and services
>
>    - You'll need to manually enter:
>
>      - Any donations/contributions (Line 1, Cell F6)
>      - Membership dues if applicable (Line 4, Cell H9)
>      - Interest/investment income if any (Line 5, Cell H10)
>      - Fundraising information if applicable (Lines 9-11)
>      - Other income if applicable (Line 12, Cell H17)
>
>    - For expenses:
>
>      - Line 15 (Cell F21): Any direct grants to California State Parks
>      - Line 16 (Cell F22): Interpretive program support expenses
>      - Line 17 (Cell F23): Other CSP program support
>      - Line 19 (Cell H25): Management and general expenses
>
>    - For assets/liabilities:
>      - Line 22 (Cell F30): Total assets
>      - Line 23 (Cell H31): Total liabilities
>      - Lines 24-26 (Cells F32-F34): Net assets by restriction type
>
> 4. **Page 4** (not fully examined but likely contains):
>    - Additional notes, explanations, and signatures
>
> For your specific activities (book publishing and firewood sales), focus on accurately completing:
>
> 1. The Interpretive and Educational Sales section on Page 2(A) for your book sales
> 2. The Firewood Sales section on Page 2(A) for your firewood revenue
> 3. The complete financial summary on Page 3
>
> Would you like me to explain any specific section in more detail?

### Example 4

This uses the file `2025-3-7-iolp-buildings.xlsx` from the Inventory of Owned and Leased Properties (IOLP) dataset available at https://catalog.data.gov/dataset/inventory-of-owned-and-leased-properties-iolp If you've followed the news in recent weeks you will know why this is currently the most popular dataset on data.gov.

Prompt:

> Using the Excel file in "/Users/jonas/Downloads/2025-3-7-iolp-buildings.xlsx": What is the largest federal building by square footage in the state of Utah?

I used this file and prompt repeatedly as test case and got mixed results:

- Chats that terminated without a response because they had reached the context limit. This happened when Claude got into a loop of querying many rows one at a time.
- Incorrect answers because Claude made "bad assumptions" along the way. For example in [this chat](https://claude.ai/share/bc919595-7df7-412f-b754-2cb67cab6b73) it decides that the largest building must be in Salt Lake City and starts to search the spreadsheet for arbitrary terms.
- Correct answers, for example [this chat](https://claude.ai/share/bae7c6d8-6541-4216-953a-5b7abecde34c): The IRS Service Center in Ogden is the largest federally-owned building in Utah on the list.

Footnote: The IRS Service Center is correct insofar that it is the largest building on the list. The "Bumblehive" aka "[Utah Data Center](https://en.wikipedia.org/wiki/Utah_Data_Center)" aka "Intelligence Community Comprehensive National Cybersecurity Initiative Data Center" is almost certainly larger but is not included in the file.
