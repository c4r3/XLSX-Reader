package com.care.ssm

import com.care.ssm.handlers.SheetHandler.{Cell, CellType, Row}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ReadingTest extends AnyFlatSpec with Matchers {

  "Reading sample_1" should "be ok" in {

    val path = "./src/test/resources/sample_1/sample.xlsx"
    val sheet = "sheet1"

    val parser = new DocumentSaxParser
    val result: List[Row] = parser.readSheet(path, sheet)

    result.size should be (391)

    val header = result.head
    header.cells.size should be (5)
    header.cells.head should be (Cell(1,1,"Postcode", CellType.String, null))
    header.cells(1) should be (Cell(1,2,"Sales_Rep_ID", CellType.String, null))
    header.cells(2) should be (Cell(1,3,"Sales_Rep_Name", CellType.String, null))
    header.cells(3) should be (Cell(1,4,"Year", CellType.String, null))
    header.cells(4) should be (Cell(1,5,"Value", CellType.String, null))

    val first = result(1)
    first.cells.head should be (Cell(2,1,2121, CellType.Integer, null))
    first.cells(1) should be (Cell(2,2,456, CellType.Integer, null))
    first.cells(2) should be (Cell(2,3,"Jane", CellType.String, null))
    first.cells(3) should be (Cell(2,4,2011, CellType.Integer, null))
    first.cells(4) should be (Cell(2,5,84219.497310686638, CellType.Currency, Map("sign" -> "$")))

    val second = result(2)
    second.cells.head should be (Cell(3,1,2092, CellType.Integer, null))
    second.cells(1) should be (Cell(3,2,789, CellType.Integer, null))
    second.cells(2) should be (Cell(3,3,"Ashish", CellType.String, null))
    second.cells(3) should be (Cell(3,4,2012, CellType.Integer, null))
    second.cells(4) should be (Cell(3,5,28322.19226785212, CellType.Currency, Map("sign" -> "$")))

    val last = result(390)
    last.cells.head should be (Cell(391,1,2116, CellType.Integer, null))
    last.cells(1) should be (Cell(391,2,456, CellType.Integer, null))
    last.cells(2) should be (Cell(391,3,"Jane", CellType.String, null))
    last.cells(3) should be (Cell(391,4,2013, CellType.Integer, null))
    last.cells(4) should be (Cell(391,5,3195.699054497647, CellType.Currency, Map("sign" -> "$")))
  }

  "Reading Doubles" should "be ok" in {

    val path = "./src/test/resources/doubles/doubles.xlsx"
    val sheet = "Foglio1"

    val parser = new DocumentSaxParser
    val result: List[Row] = parser.readSheet(path, sheet)

    result.size should be (12)

    val header = result.head
    header.cells.size should be (9)
    header.cells.head should be (Cell(1,1,"val_1", CellType.String, null))
    header.cells(1) should be (Cell(1,2,"val_2", CellType.String, null))
    header.cells(2) should be (Cell(1,3,"val_3", CellType.String, null))
    header.cells(3) should be (Cell(1,4,"val_4", CellType.String, null))
    header.cells(4) should be (Cell(1,5,"val_5", CellType.String, null))
    header.cells(5) should be (Cell(1,6,"val_6", CellType.String, null))
    header.cells(6) should be (Cell(1,7,"val_7", CellType.String, null))
    header.cells(7) should be (Cell(1,8,"val_8", CellType.String, null))
    header.cells(8) should be (Cell(1,9,"val_9", CellType.String, null))

    val first = result(1)
    first.cells.size should be (9)
    first.cells.head should be (Cell(2,1,2.0, CellType.Double, null))
    first.cells(1) should be (Cell(2,2,3.0, CellType.Double, null))
    first.cells(2) should be (Cell(2,3,4.0, CellType.Double, null))
    first.cells(3) should be (Cell(2,4,5.0, CellType.Currency,Map("sign" -> "€")))
    first.cells(4) should be (Cell(2,5,6.0, CellType.Double, null))
    first.cells(5) should be (Cell(2,6,7.0, CellType.Double, null))
    first.cells(6) should be (Cell(2,7,8, CellType.Integer, null))
    first.cells(7) should be (Cell(2,8,9.0, CellType.Double, null))
    first.cells(8) should be (Cell(2,9,1000000.1, CellType.Double, null))

    val second = result(2)
    second.cells.size should be (9)
    second.cells.head should be (Cell(3,1,2.1, CellType.Double, null))
    second.cells(1) should be (Cell(3,2,3.1, CellType.Double, null))
    second.cells(2) should be (Cell(3,3,4.1, CellType.Double, null))
    second.cells(3) should be (Cell(3,4,5.1, CellType.Currency,Map("sign" -> "€")))
    second.cells(4) should be (Cell(3,5,6.1, CellType.Double, null))
    second.cells(5) should be (Cell(3,6,7.1, CellType.Double, null))
    second.cells(6) should be (Cell(3,7,8.1, CellType.Double, null))
    second.cells(7) should be (Cell(3,8,9.1, CellType.Double, null))
    second.cells(8) should be (Cell(3,9,1000000.2, CellType.Double, null))

    val last = result(11)
    last.cells.size should be (9)
    last.cells.head should be (Cell(12,1,2.1, CellType.Double, null))
    last.cells(1) should be (Cell(12,2,3.10, CellType.Double, null))
    last.cells(2) should be (Cell(12,3,4.10, CellType.Double, null))
    last.cells(3) should be (Cell(12,4,5.10, CellType.Currency,Map("sign" -> "€")))
    last.cells(4) should be (Cell(12,5,6.10, CellType.Double, null))
    last.cells(5) should be (Cell(12,6,7.10, CellType.Double, null))
    last.cells(6) should be (Cell(12,7,9, CellType.Integer, null))
    last.cells(7) should be (Cell(12,8,10.0, CellType.Double, null))
    last.cells(8) should be (Cell(12,9,1000000.11, CellType.Double, null))
  }

  "Reading Sample_2" should "be ok" in {

    val path = "./src/test/resources/sample_2/sample_2.xlsx"
    val sheet = "mix"

    val parser = new DocumentSaxParser
    val result: List[Row] = parser.readSheet(path, sheet)

    result.size should be(15)

    val header = result.head
    header.cells.size should be(11)
    header.cells.head should be(Cell(1, 1, "integer_general", CellType.String, null))
    header.cells(1) should be(Cell(1, 2, "integer_numeric", CellType.String, null))
    header.cells(2) should be(Cell(1, 3, "contab.", CellType.String, null))
    header.cells(3) should be(Cell(1, 4, "val.", CellType.String, null))
    header.cells(4) should be(Cell(1, 5, "time", CellType.String, null))
    header.cells(5) should be(Cell(1, 6, "date", CellType.String, null))
    header.cells(6) should be(Cell(1, 7, "data_ext", CellType.String, null))
    header.cells(7) should be(Cell(1, 8, "perc.", CellType.String, null))
    header.cells(8) should be(Cell(1, 9, "date.abbr.", CellType.String, null))
    header.cells(9) should be(Cell(1, 10, "scientific", CellType.String, null))
    header.cells(10) should be(Cell(1, 11, "text", CellType.String, null))

    val first = result(1)
    first.cells.size should be(11)
    first.cells.head should be(Cell(2, 1, 1, CellType.Integer, null))
    first.cells(1) should be(Cell(2, 2, 1.0, CellType.Double, null))
    first.cells(2) should be(Cell(2, 3, 100.0, CellType.Currency, Map("sign" -> "€")))
    first.cells(3) should be(Cell(2, 4, 200.0, CellType.Currency, Map("sign" -> "€")))
    first.cells(4) should be(Cell(2, 5, 83100000, CellType.Time, null))
    first.cells(5) should be(Cell(2, 6, 1607731200000L, CellType.Date, null))
    first.cells(6) should be(Cell(2, 7, 631201500000L, CellType.Date, null))
    first.cells(7) should be(Cell(2, 8, 0.17, CellType.Double, null))
    first.cells(8) should be(Cell(2, 9, 1593907200000L, CellType.Date, null))
    first.cells(9) should be(Cell(2, 10, 1.0E8, CellType.Double, null))
    first.cells(10) should be(Cell(2, 11, "abc", CellType.String, null))
  }
}
