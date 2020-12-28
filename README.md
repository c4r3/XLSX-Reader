# XLSX-Reader
XLSX-Reader
Just a pure Scala XLSX reader. It's aimed to the parsing of XLSX file without any other libraries (i.e. no POI Stuff).

Base features:

- Read directly from zipstream
- Parsing all major excel datatypes
- Parsing single cells with metadata (ex.: currency meta)
- It can't manage formulas and excel's errors
- Range rows reading
- Date are normalized from 01-01-1970
