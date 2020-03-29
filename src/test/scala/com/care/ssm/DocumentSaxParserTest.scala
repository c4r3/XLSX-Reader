package com.care.ssm

import org.scalatest.PrivateMethodTester
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class DocumentSaxParserTest extends AnyFlatSpec with PrivateMethodTester with Matchers {

  "inner zip sheet file path" should " be correct " in {

    val parser = new DocumentSaxParser
    val buildInnerZipSheetFilePath: PrivateMethod[String] = PrivateMethod[String](Symbol("buildInnerZipSheetFilePath"))

    "xl/worksheets/sheet1.xml" should equal (parser invokePrivate buildInnerZipSheetFilePath("xl/worksheets", "1"))
    "xl/worksheets/sheet123.xml" should equal (parser invokePrivate buildInnerZipSheetFilePath("xl/worksheets", "123"))
  }
}