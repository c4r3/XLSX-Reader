import java.util

import com.care.ssm.handlers.SheetHandler
import com.care.ssm.DocumentSaxParser

import scala.collection.mutable.ListBuffer

object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser()

    val result: ListBuffer[SheetHandler.SSRawCell] = parser.readSheet(path, "sheet1", 0, 7)
    println(s"Total Cells: ${result.length}")
    result.toList.foreach(c => println(c))

    println("\n\n\n")

    val resultList = parser.lookupValues(path, result)
    println(s"Total Cells: ${resultList.length}")
    resultList.foreach(c => println(c))
  }
}