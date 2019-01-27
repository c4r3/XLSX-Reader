import com.care.ssm.{DocumentSaxParser, SSMUtils}
import javax.xml.parsers.SAXParser

object Main {

  def main(args: Array[String]): Unit = {

    val path = "./src/test/resources/sample.xlsx"

/* [Content_Types].xml
    _rels/.rels
    xl/_rels/workbook.xml.rels
    xl/workbook.xml
    xl/sharedStrings.xml
    xl/worksheets/_rels/sheet1.xml.rels
    xl/theme/theme1.xml
    xl/styles.xml
    xl/worksheets/sheet1.xml
    docProps/core.xml
    xl/printerSettings/printerSettings1.bin
    docProps/app.xml*/

/*    println("Starting Spread Sheet Manager...")
    val path = "./src/test/resources/sample.xlsx"
    val result = SSMUtils.extractStream(path,"xl/worksheets/sheet1.xml")
    println(result.getOrElse("Something wrong"))*/

/*    println("Starting Spread Sheet Manager...")
    val result = SSMUtils.extractStream(path,"xl/workbook.xml")
    println(result.getOrElse("Something wrong"))*/

    val parser = new DocumentSaxParser()

    //val result = parser.lookupSheetIdByName(path, "Sheet1")
    //val result = parser.bubu(path)

    val result = parser.lookupSharedString(path)
    println(s"Result: $result")

    val result2 = parser.lookupSharedString(path, Set[Int](1,2,3))
    println(s"Result: $result2")

    val result3 = parser.lookupSharedString(path, Set[Int](4,5,7))
    println(s"Result: $result3")

    /*ora aggiungi lo sheet reader, serve una case class per le celle e un metodo di lettura intelligente
    occorre poi un metodo in grado di fare la detection del tipo
    e agganciare il lookup alla shared strings*/

    val result5 = parser.readSheet(path, "sheet1")
    println(s"Total Cells: ${result5.size()}")

    result5.forEach(println)
  }
}