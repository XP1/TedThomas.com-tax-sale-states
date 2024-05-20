# Tax lien certificate states and tax deed states

This script fetches the data from the following lists on [TedThomas.com](https://tedthomas.com/), parses the webpages, and writes the data in Excel workbook (XLSX), CSV, JSON, and markdown table format.

* The Essential List of Tax Lien Certificate States:
    * https://tedthomas.com/faqs/tax-lien-certificate-states/


* Ted Thomas Complete List of Tax Deed States:
    * https://tedthomas.com/faqs/tax-deed-states/

See the ["build" directory](/build) for the output.

## Requirements

* [Python 3.10+](https://www.python.org/downloads/)

## Installation

```shell
pip install requests beautifulsoup4 py-markdown-table openpyxl
```

## Usage

```shell
python tax_sale_states.py
```

### Example

```
python tax_sale_states.py

Building "Tax lien certificate states"...
    Fetching data... Done.
    Writing data... Done.
    Creating workbook... Done.
    Writing workbook... Done.
    Writing CSV... Done.
    Writing JSON... Done.
    Writing markdown... Done.

Building "Tax deed states"...
    Fetching data... Done.
    Writing data... Done.
    Creating workbook... Done.
    Writing workbook... Done.
    Writing CSV... Done.
    Writing JSON... Done.
    Writing markdown... Done.
```

The ["build" directory](/build) contains the written files.

The ["data" directory](/data) contains the HTML source files.