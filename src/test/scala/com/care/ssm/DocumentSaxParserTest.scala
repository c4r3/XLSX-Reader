package com.care.ssm

import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.scalatest.PrivateMethodTester
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

import scala.collection.mutable.ListBuffer

class DocumentSaxParserTest extends AnyFlatSpec with PrivateMethodTester with Matchers {

  "inner zip sheet file path " should " be correct " in {

    val parser = new DocumentSaxParser
    val method: PrivateMethod[String] = PrivateMethod[String](Symbol("buildInnerZipSheetFilePath"))

    "xl/worksheets/sheet1.xml" should equal (parser invokePrivate method("xl/worksheets", "1"))
    "xl/worksheets/sheet123.xml" should equal (parser invokePrivate method("xl/worksheets", "123"))
  }

  "lookup cells style " should " be correct " in {

    val parser = new DocumentSaxParser
    val method = PrivateMethod[ListBuffer[SSCellStyle]](Symbol("lookupCellsStyles"))
    val result: ListBuffer[SSCellStyle] = parser invokePrivate method("./src/test/resources/sample.xlsx")

    result should have size 5
    result.head should be (null)
    result(1) should not be null
    result(1) should be (SSCellStyle(164, "\"$\"#,##0"))
    result(2) should be (null)
    result(3) should be (null)
    result(4) should not be null
    result(4) should be (SSCellStyle(1, "\"$\"#,##0"))
  }
}