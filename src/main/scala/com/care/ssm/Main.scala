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

    parser.lookupSheetIdByName(path, "Sheet1")

  }
}