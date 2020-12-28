package com.c4r3.xlsx.reader

import com.c4r3.xlsx.reader.handlers.SheetHandler.Row
import org.slf4j
import org.slf4j.LoggerFactory

/**
  * @author C4r3
  */
object ReaderMain {

  val logger: slf4j.Logger = LoggerFactory.getLogger(ReaderMain.getClass)

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample_1/sample.xlsx"
    val sheet = "sheet1"

    logger.debug("Starting parsing XLSX at path {}", path)

    val parser = new XLSXParser
    val result: List[Row] = parser.readSheet(path, sheet, 0, 5)
    result.foreach(println)
  }
}