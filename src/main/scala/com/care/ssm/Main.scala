import java.util

import com.care.ssm.handlers.SheetHandler
import com.care.ssm.{DocumentSaxParser}

object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser()

    val result = parser.lookupSharedString(path)
    println(s"Result: $result")

    val result2 = parser.lookupSharedString(path, Set[Int](1,2,3))
    println(s"Result: $result2")

    val result3 = parser.lookupSharedString(path, Set[Int](4,5,7))
    println(s"Result: $result3")

    println("\n\n\n")

    val result5: util.ArrayList[SheetHandler.SSCell] = parser.readSheet(path, "sheet1", 1, 2)
    println(s"Total Cells: ${result5.size()}")
    result5.forEach(println)
  }
}