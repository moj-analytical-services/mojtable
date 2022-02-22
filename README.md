# mojtable

> Create statistical tables that meet accessibility requirements

## Overview

This package contains functions to create statistical tables that meet accessibility requirements.

For more information on accessibility, please see: https://gss.civilservice.gov.uk/policy-store/accessibility-legislation-what-you-need-to-know/

The MoJ PDC have created a guidance document for making accessible spreadsheets here: https://moj-analytical-services.github.io/PDC-guidance/. This package is intended to make that process easier by enabling the output of an accessible spreadsheet directly from R. 

### Installation:

```r
devtools::install_github("moj-analytical-services/mojtable")
```

## Available functions

- [makeExcel](#makeExcel)

### makeExcel

Creates an Excel workbook (.xlsx) containing statistical tables that meet accessibility requirements.

The structure of each table should be created within an R data frame which is then exported as a table to an Excel worksheet. Column names in the data frame will become column names in the final Excel table.

This function accepts a series of lists as arguments. Each list contains a series of values that are used to construct the individual worksheets within the overall workbook. Each list should be the same length and arranged in the order that you want the tabs to appear in the final Excel workbook.

eg. tables[[1]] contains the table for the first tab of the workbook, titles[[1]] contains the title for this table, and so on.

> **NOTE:** This function will output a .XLSX file. You must open this file in Excel and save it in .ODS format before publishing on Gov.Uk

#### Syntax

```r
makeExcel(
  tables,
  titles,
  subtitles,
  tabnames,
  tablenames,
  filename
  )
```

#### Arguments

- **tables**: A list of data frames. The data frames will be exported to an Excel worksheet in the same order as they appear in the list. (ie. tables[[1]] will appear in the first tab of the workbook, table[[2]] in the second, and so on)

- **titles**: A list of character vectors of length 1. Each list element should contain a single string which will be displayed as the corresponding title of each table.

- **subtitles**: A list of character vectors of any length. Each list element should contain a subtitle for the corresponding table. If the vector length of any list element is > 1 then each string within the vector will be displayed on a separate row, allowing for the creation of multi-line subtitles.

- **tabnames**: A list of character vectors of length 1. Each list element should contain a name for the corresponding tab of the workbook. Tab names should be short but descriptive and can only contain alphanumeric characters and underscores.

- **tablenames**: A list of character vectors of length 1. Each list element should contain a name for the corresponding table. Table names should be short but descriptive and can only contain alphanumeric characters and underscores.

- **filename**: A string showing the filename to give to the exported Excel workbook. The filename should include the extension .xlsx
