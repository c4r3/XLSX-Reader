package com.care.ssm

import com.care.ssm.DocumentSaxParser.SSMCell
import com.care.ssm.handlers.SheetHandler.SSRawCell

import scala.collection.mutable.ListBuffer

/**
  * @author Massimo Caresana
  */
object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser
    val result: ListBuffer[SSRawCell] = parser.readSheet(path, "sheet1")
    val resultList: List[SSMCell] = parser.parseRawCells(path, result)

    println(s"Total Cells: ${resultList.length}")
    resultList.foreach(c => println(c))
  }
}