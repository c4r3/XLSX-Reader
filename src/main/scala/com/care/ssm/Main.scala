package com.care.ssm

import com.care.ssm.DocumentSaxParser.SSMCell
import com.care.ssm.handlers.SheetHandler.SSRawCell


/**
  * @author Massimo Caresana
  */
object Main {

  def main(args: Array[String]): Unit = {

    //val path = "./src/test/resources/sample_1/sample.xlsx"
    val path = "./src/test/resources/doubles/doubles.xlsx"
    val parser = new DocumentSaxParser
    //val result: ListBuffer[SSRawCell] = parser.readSheet(path, "sheet1")
    val result: List[SSRawCell] = parser.readSheet(path, "Foglio1", 0, 5)
    result.foreach(println)


    val resultList: List[SSMCell] = parser.parseRawCells(path, result)

    println(s"Total Cells: ${resultList.length}")
    resultList.foreach(c => println(c))
  }
}