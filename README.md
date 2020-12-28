# XLSX-Reader
XLSX-Reader
Just a pure Scala XLSX reader. It's aimed at parsing of XLSX file without any other libraries (i.e. no POI Stuff here).

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
