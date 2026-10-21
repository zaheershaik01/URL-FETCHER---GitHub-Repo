# Government and NIC Website URL Extractor

A Python script that downloads a domain list, extracts Indian government (`.gov.in`) and National Informatics Centre (`.nic.in`) domain names, and exports the resulting HTTPS URLs to text and Excel files.

## Features

- Downloads the source domain list from a GitHub Gist.
- Uses a regular expression to find `.gov.in` and `.nic.in` hostnames.
- Removes duplicates and sorts the results.
- Creates separate output files for each suffix:
  - `gov_in_urls.txt` and `gov_in_urls.xlsx`
  - `nic_in_urls.txt` and `nic_in_urls.xlsx`
- Adds serial numbers and a `Website URL` column to the Excel files.
- Formats the Excel header in bold, freezes the header row, and sets the URL column width.
- Retries a failed download up to three times, waiting five seconds between attempts.

## Requirements

- Python 3.9 or later (recommended)
- `requests`
- `openpyxl`

## Installation

1. Save `script1_fetch_all_urls.py` in your project directory.
2. (Optional but recommended) Create and activate a virtual environment.
3. Install dependencies:

   ```bash
   pip install requests openpyxl
   ```

## Usage

Run the script from a terminal in the project directory:

```bash
python script1_fetch_all_urls.py
```

The script needs an internet connection to download the source list.

## Output

When matching domains are found, the script writes files to the current working directory.

| File | Contents |
|---|---|
| `gov_in_urls.txt` | Sorted HTTPS URLs for domains ending in `.gov.in` |
| `gov_in_urls.xlsx` | The same URLs in an Excel worksheet |
| `nic_in_urls.txt` | Sorted HTTPS URLs for domains ending in `.nic.in` |
| `nic_in_urls.xlsx` | The same URLs in an Excel worksheet |

If no matching domains are found, the script prints a preview of the downloaded text and exits without writing files. If no domains are found for one suffix, files for that suffix are skipped.

## Error Handling

- **Download failure:** Retries up to three times, then exits with an error message.
- **Excel file is open or locked:** Prints a message asking you to close the workbook and run the script again.
- **No matching domains:** Prints the first 300 characters of the downloaded content and stops.

## Notes

- The script extracts domain names from the downloaded list; it does **not** check whether each website is reachable.
- The output URLs are constructed with `https://`.
- The source list is hosted at a GitHub Gist. Its availability and contents may change.
