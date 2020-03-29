package com.care.ssm

import org.scalatest.PrivateMethodTester
import org.scalatest.flatspec.AnyFlatSpec

class DocumentSaxParserTest extends AnyFlatSpec with PrivateMethodTester{

  "inner zip sheet file path" should " be correct " in {

    val parser = new DocumentSaxParser
    val buildInnerZipSheetFilePath: PrivateMethod[String] = PrivateMethod[String](Symbol("buildInnerZipSheetFilePath"))

    assertResult("xl/worksheets/sheet1.xml") {
      parser invokePrivate buildInnerZipSheetFilePath("xl/worksheets", "1")
    }

    assertResult("xl/worksheets/sheet123.xml") {
      parser invokePrivate buildInnerZipSheetFilePath("xl/worksheets", "123")
    }
  }
}