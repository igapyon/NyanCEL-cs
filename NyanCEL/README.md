# NyanCEL

NyanCEL is a library that allows you to query the contents of Excel workbooks (.xlsx) using SQL.

- NyanCEL-cs is implemented in C#.
- By providing an Excel workbook (.xlsx), each sheet is treated as a table that can be queried with SQL.
- It is released under the MIT license.
- NyanCEL-cs is part of the NyanCEL project.

## Relationship with the NyanQL Project

- NyanCEL is a friend project of the Nyankuru project.
- It is an independent project inspired by the Nyankuru project.
- NyanCEL respects and honors the Nyankuru project.

## Workflow

1. Provide an Excel workbook (.xlsx) file
   - Use the sheet name as the table name
   - Use the values in the first row as column names
   - Derive column data types from the format of the second row
2. Load the contents of the specified Excel workbook into an SQLite database
   - Read data from the second row onward
   - Load the read data into an in-memory SQLite database
3. Execute a SELECT statement
   - Query the database with the provided SELECT statement
   - SQLite SQL syntax is available
4. Return the results of the SELECT statement
   - Return the search results row by row
   - Default returns the search results as JSON data
   - Change the return format to XML by adding the `fmt=xml` parameter
   - Change the return format to xlsx by adding the `fmt=xlsx` parameter
   - Narrow down return data by specifying `fmt=json&target=data.1`
   - Apply jsonpath to search results with `fmt=json&jsonpath=`
   - Apply xpath to search results with `fmt=xml&xpath=`

## Internally Used OSS

The following OSS is used internally. Thanks to the providers of each OSS.

- ClosedXML
  - MIT
  - Version 0.102.2
- Microsoft.Data.Sqlite
  - MIT
  - Version 8.0.6
- Newtonsoft.Json
  - MIT
  - Version 13.0.3

## Limitations

- It operates as an in-memory RDBMS, so it may not work with large amounts of data
- Supports .xlsx files
- Double quotes cannot be included in Excel sheet names or column names in the title row

## Excel to SQLite Type Mapping

Internally, cell value types are derived from Excel formats to SQLite types.

| Excel Format | SQLite Type | Example |
|--------------|-------------|---------|
| 0            | TEXT        |         |
| 1            | INTEGER     | 0       |
| 2            | REAL        | 0.00    |
| 3            | REAL        | #,##0   |
| 4            | REAL        | #,##0.00|
| 9            | REAL        | 0%      |
| 10           | REAL        | 0.00%   |
| 11           | TEXT        | 0.00E+00|
| 12           | TEXT        | # ?/?   |
| 13           | TEXT        | # ??/?? |
| 14           | TEXT        | d/m/yyyy: Format yyyy-MM-dd |
| 15           | TEXT        | d-mmm-yy: Format yyyy-MM-dd |
| 16           | TEXT        | d-mmm: Format MM-dd |
| 17           | TEXT        | mmm-yy: Format yyyy-MM |
| 18           | TEXT        | h:mm tt |
| 19           | TEXT        | h:mm:ss tt |
| 20           | TEXT        | H:mm: Format HH:mm |
| 21           | TEXT        | H:mm:ss: Format HH:mm:ss |
| 22           | TEXT        | m/d/yyyy H:mm: Format yyyy-MM-dd HH:mm |
| 37           | INTEGER     | #,##0;(#,##0) |
| 38           | INTEGER     | #,##0;[Red](#,##0) |
| 39           | REAL        | #,##0.00;(#,##0.00) |
| 40           | REAL        | #,##0.00;[Red](#,##0.00) |
| 45           | TEXT        | mm:ss: Format mm:ss |
| 46           | TEXT        | [h]:mm:ss: Format HH:mm:ss |
| 47           | TEXT        | mmss.0 |
| 48           | TEXT        | ##0.0E+0 |
| 49           | TEXT        |