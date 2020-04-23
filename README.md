# `copy_xlsx_styles` - A script to copy spreadsheet styles to new data batches

Command-line and Python tool to copy Excel styles from an exemplar worksheet to a dumped data worksheet.

## Purpose

Automate the styling of data in spreadsheets without programming.

We often look at data in spreadsheets. But we spend the first while styling
them (changing column widths, bolding headers, freezing panes, adding
filtering) so that they are navigable. When data is generated in repeated
batches, this styling should be automated. But that automation can happen by
copying from an example styled spreadsheet rather than programming each style.

This tool is designed to enable the following workflow for presenting data as
an Excel spreadsheet:

1. Export some data (from an online resource or a data analysis) as a CSV or
   spreadsheet.
2. Style the spreadsheet manually (including conditional formatting, data
   validation, etc.) and save this exemplar.
3. Export a new batch of data.
4. Apply the spreadsheet styles using `copy_xlsx_styles`.
5. Deliver the data spreadsheet to the recipient.

This makes it easy to have automatic processes to style a spreadsheet, but
avoids the effort of programmatically styling the spreadsheet using a tool like:
* [`styleframe`](https://styleframe.readthedocs.io),
* [`pandas.DataFrame.style`](https://pandas.pydata.org/pandas-docs/stable/user_guide/style.html),
* [`openpyxl`](https://openpyxl.readthedocs.io),
* [`xlsxwriter`](https://xlsxwriter.readthedocs.io)

## Usage

Install the `Python 3` and `openpyxl` dependencies. Basic usage is:

```
usage: copy_xlsx_styles.py [-h] style_worksheet data_worksheet output_xlsx
```

Here:
* `style_worksheet` and `data_worksheet` may be the path to an xlsx file, or
the path may have a suffix `!SheetName` to specify a specific sheet. E.g. `/path/to/spreadsheet.xlsx` (first sheet will be used) or `/path/to/spreadsheet.xlsx!Sheet2`
* `data_worksheet` may instead be a path to a CSV file: `/path/to/data.csv`
* `output_xlsx` should be the path to a `.xlsx` file to dump the result.


Within Python, the `copy_styles` function may be used directly to port styles
between two `openpyxl.WorkSheet` objects.

## Developed by the Sydney Informatics Hub

This tool was developed by the Sydney Informatics Hub, a core research facility
of The University of Sydney.

Please acknowledge the Sydney Informatics Hub in publications where the tool is
useful to you.

         /  /\        ___          /__/\   
        /  /:/_      /  /\         \  \:\  
       /  /:/ /\    /  /:/          \__\:\ 
      /  /:/ /::\  /__/::\      ___ /  /::\
     /__/:/ /:/\:\ \__\/\:\__  /__/\  /:/\:\
    \  \:\/:/~/:/    \  \:\/\ \  \:\/:/__\/
      \  \::/ /:/      \__\::/  \  \::/    
       \__\/ /:/       /__/:/    \  \:\    
         /__/:/ please \__\/      \  \:\   
         \__\/ acknowledge your use\__\/   

