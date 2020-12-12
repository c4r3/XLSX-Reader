package com.care.ssm.handlers

import java.lang.Integer._

import com.care.ssm.SSMUtils
import com.care.ssm.SSMUtils.{SSCellType, extractStream, shared_strings, toDouble, toInt}
import com.care.ssm.handlers.SheetHandler.{Cell, Row}
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import javax.xml.parsers.{SAXParser, SAXParserFactory}
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer


/**
  *
  * @author Massimo Caresana
  *
  * Handler for Shared Strings File
  * @param fromRow Start reading from this zero-based index
  * @param toRow End reading to this zero-based index
  * @param stylesList The style data
  */
class SheetHandler(fromRow: Int = 0, toRow: Int = MAX_VALUE, stylesList: List[SSCellStyle], xlsxPath: String) extends DefaultHandler{

  var result: ListBuffer[Row] = ListBuffer[Row]()

  //Child Elements
  val cellTag = "c"
  val valueTag = "v"
  val rowTag = "row"
  val formula = "f"
  val richTextInline = "is"
  val futureFeatureDataStorageArea = "extLst"

  //Attributes
  val rowNumAttr = "r"
  val xyAttr = "r"
  val styleAttr = "s"
  val typeAttr = "t"
  val cellMetadataIndexAttr = "cm"
  val showPhoneticAttr = "ph"
  val valueMetadataIndexAttr = "vm"

  //Flags
  var cellEnded = false
  var valueTagStarted = false

  //Temporary variables
  var rowNum: Int = -1
  var cellXY = ""
  var cellStyleIndex: Int = -1
  var cellType: String = null

  var rowCellsBuffer: ListBuffer[Cell] = ListBuffer[Cell]()

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    //TODO controllare, serve solo uno dei due
    if (workDone || isNotRequiredRow) return

    if(formula.equals(qName)) {
      print("warning: no formula is supported, check the workbook")
      return
    }

    if(richTextInline.equals(qName)) {
      print("warning: no rich text inline is supported, check the workbook")
      return
    }

    if(richTextInline.equals(qName)) {
      print("warning: no future feature data storage area supported, check the workbook")
      return
    }

    if(cellMetadataIndexAttr.equals(qName)) {
      print("warning: no cell metadata index attribute area supported, check the workbook")
      return
    }

    if(showPhoneticAttr.equals(qName)) {
      print("warning: no phonetic attribute area supported, check the workbook")
      return
    }

    if(valueMetadataIndexAttr.equals(qName)) {
      print("warning: no value metadata index attribute area supported, check the workbook")
      return
    }

    if(rowTag.equals(qName)) {
      val rawRowNumValue = attributes.getValue(rowNumAttr)
      rowNum = valueOf(rawRowNumValue)
    }

    if(valueTag.equals(qName)){
      valueTagStarted = true
    }

    if(cellTag.equals(qName)) {
      //Starting "c" tag, extraction of the attributes
      cellXY = attributes.getValue(xyAttr)

      val styleAttrIndex = attributes.getValue(styleAttr)
      cellStyleIndex = if(styleAttrIndex!=null) {
        SSMUtils.toInt(attributes.getValue(styleAttr)).getOrElse(-1)
      } else {
        -1
      }

      cellType = attributes.getValue(typeAttr)
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {
    //if (workDone || isNotRequiredRow) return

    if(rowTag.equals(qName)) {
      result += Row(rowNum, rowCellsBuffer.toList)

      //reset temp stuff
      rowCellsBuffer.clear()
      rowNum = -1
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone || isNotRequiredRow) return

    if(valueTagStarted) {

      val style: String = if(cellStyleIndex>0) {
        stylesList(cellStyleIndex).formatCode
      } else {
        null
      }

      //result += SSRawCell(cellRowNum, calculateColumn(cellXY, cellRowNum), cellXY, cellType, new String(ch, start, length), style)
      rowCellsBuffer += Cell(evaluate(cellType, new String(ch, start, length), style))

      //Flushing buffer & reset temporary stuff
      valueTagStarted = false
      cellXY = ""
      cellStyleIndex = -1
      cellType = null
    }
  }

  def evaluate(typeString: String, stringValue: String, style: String): Any = {

    typeString match {
      //TODO da correggere, bisogna sistemare tutte le casistiche(vedi sotto per referenza in PDF)
      case "d" => applyStyle(stringValue, style, isSharedString = true)
      case "e" => null //TODO da completare
      case "inlineStr" => applyStyle(stringValue, style, isSharedString = false)
      case "s" => applyStyle(stringValue, style, isSharedString = true)
      case "n" => applyStyle(stringValue, style, isSharedString = false)
      case "b" => applyStyle(stringValue, style, isSharedString = false)
      case "str" => applyStyle(stringValue, style, isSharedString = false)
      case null =>  applyStyle(stringValue, style, isSharedString = false)
      case _ => null
    }
  }

  def lookupSharedString(ids: Set[Int]): ListBuffer[String] = {

    val factory: SAXParserFactory = SAXParserFactory.newInstance
    val parser: SAXParser = factory.newSAXParser
    val zis = extractStream(xlsxPath, shared_strings)
    val handler = new SharedStringsHandler(ids)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def applyStyle(rawVal: String, style: String, isSharedString: Boolean): Any = {

    //Lookup into shared string if required
    val stringValue = if(isSharedString) {
      val index = rawVal.toInt
      lookupSharedString(Set(index)).head
    } else {
      rawVal
    }

   style match {
      case "General" => stringValue
      case "0" => toInt(stringValue).getOrElse(0)
      case "0.00" => toDouble(stringValue).getOrElse(0.0)
      case "#,##0" => toDouble(stringValue).getOrElse(0.0)
      case "#,##0.00" => toDouble(stringValue).getOrElse(0.0)
      case "0%" => toDouble(stringValue).getOrElse(0.0)
      case "0.00%" => toDouble(stringValue.replace("%","")).getOrElse(0.0)
      case "0.00E+00" => toDouble(stringValue).getOrElse(0.0)
      case "#?/?" => toDouble(stringValue).getOrElse(0.0)
      case "#??/??" => toDouble(stringValue).getOrElse(0.0)
      case "mm-dd-yy" => None //FIXME completare
      case "d-mmm-yy" => None //FIXME completare
      case "d-mmm" => None //FIXME completare
      case "mmm-yy" => None //FIXME completare
      case "h:mm AM/PM" => None //FIXME completare
      case "h:mm:ss AM/PM" => None //FIXME completare
      case "h:mm" => None //FIXME completare
      case "h:mm:ss" => None //FIXME completare
      case "m/d/yy h:mm" => None //FIXME completare
      case "#,##0;(#,##0)" => toDouble(stringValue).getOrElse(0.0)
      case "#,##0 ;[Red](#,##0)" => toDouble(stringValue).getOrElse(0.0)
      case "#,##0.00;(#,##0.00)" => toDouble(stringValue).getOrElse(0.0)
      case "#,##0.00;[Red](#,##0.00)" => toDouble(stringValue).getOrElse(0.0)
      case "mm:ss" => None //FIXME completare
      case "[h]:mm:ss" => None //FIXME completare
      case "mmss.0" => None //FIXME completare
      case "##0.0E+0" => toDouble(stringValue).getOrElse(0.0)
      case "@" => stringValue

        //FIXMe questi sono custom, così non scala, deve essere parsato lo stile da un metodo ad hoc
      case """#,##0.00\ "€"""" => toDouble(stringValue).getOrElse(0.0)
      case "0.0000" => toDouble(stringValue).getOrElse(0.0)
      case null => stringValue
      case _ => {
        println(s"Unknown style $style")
        null
      }
    }
  }

  private def isNotRequiredRow: Boolean = !(rowNum<0 || (fromRow.toInt <= rowNum && rowNum <= toRow.toInt))

  private def workDone: Boolean = rowNum > toRow.toInt

  def getResult: List[Row] = result.toList
}

object SheetHandler {

  case class Row(index: Int, cells: List[Cell])

  //case class SSRawCell(row: Int, column: Int, xy: String, ctype: String, value: String, style: SSCellStyle)
  //case class Cell(row: Int, column: Int, xy: String,  value: Any)
  case class Cell(value: Any)
}