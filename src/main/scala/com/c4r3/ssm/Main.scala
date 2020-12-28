package com.c4r3.ssm

import com.c4r3.ssm.handlers.SheetHandler.Row
import com.c4r3.ssm.handlers.SheetHandler.Row
import com.c4r3.ssm.handlers.SheetHandler.Row
import org.slf4j
import org.slf4j.LoggerFactory

/**
  * @author Massimo Caresana
  */
object Main {

  val logger: slf4j.Logger = LoggerFactory.getLogger(Main.getClass)

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample_1/sample.xlsx"
    val sheet = "sheet1"

    logger.debug("Starting parsing XLSX at path {}", path)

    val parser = new DocumentSaxParser
    val result: List[Row] = parser.readSheet(path, sheet, 0, 5)
    result.foreach(println)
  }
}