[![Apache-2 License](https://img.shields.io/github/license/c4r3/XLSX-Reader)](https://github.com/c4r3/XLSX-Reader/blob/master/LICENSE.md)
[![Build Status](https://travis-ci.com/c4r3/XLSX-Reader.svg?branch=master)](https://travis-ci.com/c4r3/XLSX-Reader)
[![codecov](https://codecov.io/gh/c4r3/XLSX-Reader/branch/master/graph/badge.svg)](https://codecov.io/gh/c4r3/XLSX-Reader)

# XLSX-Reader (100% Scala Excel files reader)

Just a pure Scala XLSX reader for the famous office suite spreadsheets generator. 
It's aimed at parsing of XLSX file without any other libraries. No Apache POI dependencies has been used for this project.
The business logic is fast and lightweight; with useful API to access the data and the smallest possible amount of datastructures. The functional approach with some container entities avoid the cumbersome boilerplate of the others object oriented based libraries. 
The use of shared states is reduced at the minimum and under the hood the parsing is based on a SAX parser on zipstreams readen directly from the internal files. This architecture reduce the memory occupation and is very fast; especially compared to the old DOM based ones.
Row is the main abstraction that map the row into the sheet with just the semantic content after a sanitization and ad hoc transformation phase.

Base features:

- Read directly from zipstreams
- Parsing all major excel datatypes
- Parsing cells with metadata (ex.: currency meta)
- It can't manage formulas and excel errors
- Range rows reading
- Dates are normalized from 01-01-1970

Simple use, just create the parser instance and read the target sheet:

```
  val path = "/your/path/to/file.xlsx"
  val sheet = "Sheet_1"
  val fromRow = 0
  val toRow = 5
  val parser = new XLSXParser
  val result: List[Row] = parser.readSheet(path, sheet, fromRow, toRow)
```
