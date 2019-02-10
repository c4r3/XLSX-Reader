import java.util

import com.care.ssm.handlers.SheetHandler
import com.care.ssm.{DocumentSaxParser}

object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser()

    val result: util.ArrayList[SheetHandler.SSCell] = parser.readSheet(path, "sheet1", 0, 7)
    println(s"Total Cells: ${result.size()}")
    result.forEach(println)
  }
}