package com.care.ssm

import com.care.ssm.handlers.SheetHandler.Row


/**
  * @author Massimo Caresana
  */
object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample_1/sample.xlsx"
    val sheet = "sheet1"

    //val path = "./src/test/resources/doubles/doubles.xlsx"
    //val sheet = "Foglio1"

    val parser = new DocumentSaxParser
    val result: List[Row] = parser.readSheet(path, sheet, 0, 5)
    result.foreach(println)
  }
}