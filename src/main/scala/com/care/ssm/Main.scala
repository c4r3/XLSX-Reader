
import com.care.ssm.DocumentSaxParser


object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"
    val parser = new DocumentSaxParser

    val result = parser.readSheet(path, "sheet1", 0, 3)
    result.foreach(c => println(c))
    println("\n")
    val resultList = parser.lookupValues(path, result)
    println(s"Total Cells: ${resultList.length}")
    resultList.foreach(c => println(c))
  }
}