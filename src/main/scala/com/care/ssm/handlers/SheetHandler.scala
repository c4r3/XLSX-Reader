package com.care.ssm.handlers

import java.lang.Integer._

import com.care.ssm.SSMUtils
import com.care.ssm.SSMUtils._
import com.care.ssm.handlers.SheetHandler.CellType.CellType
import com.care.ssm.handlers.SheetHandler._
import com.care.ssm.handlers.StyleHandler.CellStyle
import javax.xml.parsers.{SAXParser, SAXParserFactory}
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer
import scala.util.control.Exception.allCatch


/**
  *
  * @author Massimo Caresana
  *
  *         Handler for Shared Strings File
  * @param fromRow    Start reading from this zero-based index
  * @param toRow      End reading to this zero-based index
  * @param stylesList The style data
  */
class SheetHandler(fromRow: Int = 0, toRow: Int = MAX_VALUE, stylesList: List[CellStyle], xlsxPath: String) extends DefaultHandler {

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

    if (formula.equals(qName)) {
      print("warning: no formula is supported, check the workbook")
      return
    }

    if (richTextInline.equals(qName)) {
      print("warning: no rich text inline is supported, check the workbook")
      return
    }

    if (richTextInline.equals(qName)) {
      print("warning: no future feature data storage area supported, check the workbook")
      return
    }

    if (cellMetadataIndexAttr.equals(qName)) {
      print("warning: no cell metadata index attribute area supported, check the workbook")
      return
    }

    if (showPhoneticAttr.equals(qName)) {
      print("warning: no phonetic attribute area supported, check the workbook")
      return
    }

    if (valueMetadataIndexAttr.equals(qName)) {
      print("warning: no value metadata index attribute area supported, check the workbook")
      return
    }

    if (rowTag.equals(qName)) {
      val rawRowNumValue = attributes.getValue(rowNumAttr)
      rowNum = valueOf(rawRowNumValue)
    }

    if (valueTag.equals(qName)) {
      valueTagStarted = true
    }

    if (cellTag.equals(qName)) {
      //Starting "c" tag, extraction of the attributes
      cellXY = attributes.getValue(xyAttr)

      val styleAttrIndex = attributes.getValue(styleAttr)
      cellStyleIndex = if (styleAttrIndex != null) {
        SSMUtils.toInt(attributes.getValue(styleAttr)).getOrElse(-1)
      } else {
        -1
      }

      cellType = attributes.getValue(typeAttr)
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {
    if (workDone || isNotRequiredRow) return

    if (rowTag.equals(qName)) {
      result += Row(rowNum, rowCellsBuffer.toList)

      //reset temp stuff
      rowCellsBuffer.clear()
      rowNum = -1
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone || isNotRequiredRow) return

    if (valueTagStarted) {

      val style =
        if (cellStyleIndex > 0) {
          stylesList(cellStyleIndex)
        } else {
          null
        }

      //TODO aggiungere cellXY se serve
      val colNum = calculateColumn(cellXY, rowNum)
      val stringValue = new String(ch, start, length)
      rowCellsBuffer += evaluate(rowNum, colNum, cellType, stringValue, style)

      //Flushing buffer & reset temporary stuff
      valueTagStarted = false
      cellXY = ""
      cellStyleIndex = -1
      cellType = null
    }
  }

  /**
    *
    * https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx
    * <pre>
    * -------------------------------------------------------------------------
    * Enumeration Value          Description
    * -------------------------------------------------------------------------
    * b (Boolean)                Cell containing a boolean.
    * d (Date)                   Cell contains a date in the ISO 8601 format.
    * e (Error)                  Cell containing an error.
    * inlineStr (Inline String)  Cell containing an (inline) rich string, i.e., one not in the shared string table.
    * If this cell type is used, then the cell value is in the is element rather
    * than the v element in the cell (c element).
    * n (Number)                 Cell containing a number.
    * s (Shared String)          Cell containing a shared string.
    * str (String)               Cell containing a formula string
    * </pre>
    *
    * PDF Part 4, pag 2840, p. 3.18.12
    *
    * @return The detected SSCellType
    */
  def evaluate(rowNum: Int, colNum: Int, typeString: String, stringValue: String, style: CellStyle): Cell =

    typeString match {
      //TODO da correggere, bisogna sistemare tutte le casistiche(vedi sotto per referenza in PDF)
      case "s" | "d" => applyStyle(rowNum, colNum, stringValue, style, isSharedString = true)
      case "e" => null //TODO da completare
      case "str" => null //TODO da completare: cell contains a formula
      case "inlineStr" | "n" | "b" | null => applyStyle(rowNum, colNum, stringValue, style, isSharedString = false)
      case _ => null
    }

  def isInt(s: String): Boolean = (allCatch opt s.toInt).isDefined

  def isDouble(s: String): Boolean = (allCatch opt s.toDouble).isDefined

  def applyStyle(rowNum: Int, colNum: Int, rawVal: String, style: CellStyle, isSharedString: Boolean): Cell = {

    //Lookup into shared string if required
    val stringValue = if (isSharedString) {
      val index = rawVal.toInt
      lookupSharedString(Set(index), xlsxPath).head
    } else {
      rawVal
    }

    val formatCode = if(style!=null) {
      sanitizeFormatCode(style.formatCode)
    } else {
      null
    }

    var cellType = CellType.Unknown
    val meta = scala.collection.mutable.Map[String, Any]()

    val value = {

      if ((formatCode == null && isSharedString) || formatCode == "General" || formatCode == "@") {

        val w = if(isInt(stringValue)) {

          cellType = CellType.Integer
          toInt(stringValue).getOrElse(0)
        } else if(isDouble(stringValue)) {

          cellType = CellType.Double
          toDouble(stringValue).getOrElse(0.0)
        } else {
          cellType = CellType.String
          stringValue
        }
        w
      } else if (formatCode == null && !isSharedString) {

        val v: Any = if (stringValue.contains(".")) {
          cellType = CellType.Double
          toDouble(stringValue).getOrElse(0.0)
        } else if(isInt(stringValue)) {
          cellType = CellType.Integer
          toInt(stringValue).getOrElse(0)
        }   else {
          cellType = CellType.String
          stringValue
        }
        v
      } else if (formatCode == "0") {

        cellType = CellType.Integer
        toInt(stringValue).getOrElse(0)
      } else if (isCharsSubset(formatCode, "0.,#")) {

        cellType = CellType.Double
        toDouble(stringValue).getOrElse(0.0)
      } else if (isCharsSubset(formatCode, "0.%")) {

        cellType = CellType.Double
        toDouble(stringValue.replace("%", "")).getOrElse(0.0)
      } else if (formatCode == "0.00E+00") {

        cellType = CellType.Double
        toDouble(stringValue).getOrElse(0.0)
      } else if (isCharsSubset(formatCode, "ms:h.0")) {

        cellType = CellType.Time
        parseTime(stringValue)
      } else if (isCharsSubset(formatCode, "mdy-, /h")) {

        cellType = CellType.Date
        parseDateStringWithFormat(stringValue)
      } else if (isCharsSubset(formatCode, "\"0,.#$€ ")) {

        val first = formatCode.indexOf("\"")
        val last = formatCode.lastIndexOf("\"")
        val text = formatCode.substring(first + 1, last)

        if (text.trim.length == 1) {
          meta += ("sign" -> text.trim)
        } else {
          meta += ("text" -> text.trim)
        }

        cellType = CellType.Currency
        toDouble(stringValue).getOrElse(0.0)
      } else {

        formatCode match {

          case "#?/?" | "#??/??" =>

            val d = toDouble(stringValue)

            //TODO se non è parsabile come double lo si declassa a stringa
            val res = if(d.isDefined) {
              cellType = CellType.Double
              d
            } else {
              cellType = CellType.String
              stringValue
            }
            res
          case "##0.0E+0" =>
            cellType = CellType.Double
            toDouble(stringValue).getOrElse(0.0)
          case _ =>
            println(s"Unknown style $style")
            null
        }
      }
    }

    val metaMap = if (meta != null && meta.isEmpty) {
      null
    } else {
      meta.toMap
    }

    Cell(rowNum, colNum, value, cellType, metaMap)
  }

  def isCharsSubset(formatCode: String, chars: String): Boolean = {

    if (formatCode == null || formatCode.trim.isEmpty) return false

    val charsArray = chars.toCharArray

    for (c <- formatCode) {

      if (!charsArray.contains(c)) {
        return false
      }
    }
    true
  }

  private def isNotRequiredRow: Boolean = !(rowNum < 0 || (fromRow.toInt <= rowNum && rowNum <= toRow.toInt))

  private def workDone: Boolean = rowNum > toRow.toInt

  def getResult: List[Row] = result.toList
}

object SheetHandler {

  final val MILLIS_IN_DAY = 86_400_000L
  final val DAYS_FROM_0_1_1900 = 25_569.0

  object CellType extends Enumeration {
    type CellType = Value
    val Integer, Double, Long, String, Currency, Time, Date, Unknown = Value
  }

  case class Row(index: Int, cells: List[Cell]) {

    def toCSV: String = {
     if(cells!=null && cells.nonEmpty) {
       cells.map(_.toCSV).mkString(";")
     } else {
       ""
     }
    }
  }

  case class Cell(rowNum: Int, colNum: Int, value: Any, cellType: CellType, meta: Map[String, Any] = null) {

    def toCSV: String = {

      val metaValues = if(meta!=null && meta.nonEmpty) {
        meta.map(pair => s"${pair._1}=${pair._2}").mkString("|")
      } else {
        ""
      }

     s""""${rowNum.toString}","$colNum","$value","$cellType","$metaValues""""
    }
  }

  def lookupSharedString(ids: Set[Int], xlsxPath: String): List[String] = {

    val factory: SAXParserFactory = SAXParserFactory.newInstance
    val parser: SAXParser = factory.newSAXParser
    val zis = extractStream(xlsxPath, shared_strings)
    val handler = new SharedStringsHandler(ids)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def parseTime(timeStr: String): Long = {
    (toDouble(timeStr).getOrElse(0.0) * 86_400_000L).round
  }

  //TODO gestire gli errori di parsing
  def parseDateStringWithFormat(timeStr: String): Long = {
    ((toDouble(timeStr).getOrElse(0.0) - DAYS_FROM_0_1_1900) * MILLIS_IN_DAY).round
  }

  def sanitizeFormatCode(formatCode: String): String = {
    if (formatCode == null || formatCode.trim.isEmpty) {
      null
    } else {

      //Sanitize useless char
      val temp = formatCode
        .replace("\\", "")
        .replaceAll("-","")
        .replaceAll("\\*","")
        .replaceAll("_","")
        .replaceAll("\\[(.*?)\\]","")
        .replace("AM/PM","") //Cleaning useless data with time
        .trim

      //Rimozione eventuale inutile suffisso
      //#,##0.00;(#,##0.00) -> #,##0.00
      if (temp.contains(";")) {
        temp.substring(0, temp.indexOf(";"))
      } else {
        temp
      }
    }
  }
}