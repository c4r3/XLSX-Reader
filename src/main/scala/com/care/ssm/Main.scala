package com.care.ssm


object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser
    val result = parser.readSheet(path, "sheet1")

    val resultList = parser.lookupValues(path, result)
    println(s"Total Cells: ${resultList.length}")
    resultList.foreach(c => println(c))
  }
}