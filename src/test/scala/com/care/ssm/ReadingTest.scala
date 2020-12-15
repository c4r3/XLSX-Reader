package com.care.ssm

import com.care.ssm.handlers.SheetHandler.{Cell, CellType, Row}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class ReadingTest extends AnyFlatSpec with Matchers {

  "Reading sample_1" should "be ok" in {


    val path = "./src/test/resources/sample_1/sample.xlsx"
    val sheet = "sheet1"

    //val path = "./src/test/resources/doubles/doubles.xlsx"
    //val sheet = "Foglio1"

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

    val firstRow = result(1)
    firstRow.cells.head should be (Cell(2,1,2121, CellType.Integer, null))
    firstRow.cells(1) should be (Cell(2,2,456, CellType.Integer, null))
    firstRow.cells(2) should be (Cell(2,3,"Jane", CellType.String, null))
    firstRow.cells(3) should be (Cell(2,4,2011, CellType.Integer, null))
    firstRow.cells(4) should be (Cell(2,5,84.219497310686638, CellType.Currency, Map("sign" -> "$")))

  }
}
