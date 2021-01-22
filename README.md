# XLSX-Reader (100% Scala Excel files reader)

Just a pure Scala XLSX reader for the famous office suite spreadsheets generator. 
It's aimed at parsing of XLSX file without any other libraries. No Apache POI dependencies has been used for this project.
The business logic is fast and lightweight; with useful API to access the data and the smallest possible amount of datastructures.
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
  val parser = new XLSXParser
  val result: List[Row] = parser.readSheet(path, sheet, 0, 5)
```
