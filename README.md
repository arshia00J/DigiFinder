# Digikala Link Finder

This Python script reads a list of words from an Excel file, searches each word on Google, checks if the results contain a link to the "digikala.com" website, and saves the results in another Excel file. The search is performed only for words that meet specific criteria.

## Features

- Reads words from an Excel file.
- Uses Selenium to perform Google searches.
- Parses the search results to find links to "digikala.com".
- Saves the results in an Excel file.
- Stops reading words from the input file if an empty cell is encountered.

## Prerequisites

- Python 3.6 or higher
- Google Chrome browser installed
- Required Python libraries:
  - `openpyxl`
  - `selenium`
  - `beautifulsoup4`
  - `webdriver_manager`

## Installation

1. Clone the repository or download the script to your local machine.
2. Install the required Python libraries using pip:

```bash
pip install openpyxl selenium beautifulsoup4 webdriver_manager
```

## Usage

1. Prepare an Excel file (`words.xlsx`) with a list of words to search for. Each word should be in a separate cell in the first column. The process will stop if an empty cell is encountered.
2. Run the script with the following command:

```bash
python digikala_link_finder.py
```

3. The script will read the words from `words.xlsx`, perform Google searches, and save the results in `results.xlsx`.

## Script Details

### read_words(file_path)

Reads the list of words from the specified Excel file (`file_path`). Stops reading if an empty cell is encountered.

### search_google(driver, query)

Uses Selenium to perform a Google search for the given `query` and returns the parsed HTML content using BeautifulSoup.

### find_digikala_link(soup)

Parses the HTML content to find and return the first link that contains "digikala.com". If no such link is found, returns `'-'`.

### save_results(results, output_file)

Writes the list of results to the specified Excel file (`output_file`).

### main(input_file, output_file)

Coordinates the overall process:
- Reads the list of words from the input Excel file.
- Initializes the Selenium WebDriver with headless Chrome.
- Iterates over each word, and for words that contain any of the specified keywords (e.g., "کد", "ساختنی مدل", "گوشی", "ظرفیت", "گیگابایت", "سایز", "سانتی متر", "سانتیمتر", "ارتفاع"), performs a Google search and checks for a Digikala link.
- Saves the results in the output Excel file.
- Stops reading if an empty cell is encountered.

## Example

If the input file `words.xlsx` contains:

```
کد 12345
مدل 67890
گوشی سامسونگ
ظرفیت 500 گیگابایت
سایز 42
سانتی متر
ارتفاع 150 سانتیمتر
مدلabc
```

The script will perform Google searches for words containing the specified keywords and save the results in `results.xlsx`.

## License

This project is licensed under the MIT License.
