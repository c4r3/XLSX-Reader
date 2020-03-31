package com.care.ssm

import java.lang.Integer.MAX_VALUE
import java.sql.Date

import com.care.ssm.DocumentSaxParser._
import com.care.ssm.SSMUtils._
import com.care.ssm.handlers.SheetHandler.SSRawCell
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import com.care.ssm.handlers.{BaseHandler, SharedStringsHandler, SheetHandler, StyleHandler}
import javax.xml.parsers.SAXParserFactory

import scala.collection.mutable.ListBuffer


/**
  * @author Massimo Caresana
  */
class DocumentSaxParser {

  val factory = SAXParserFactory.newInstance
  val parser = factory.newSAXParser

  /**
    * Look up sheet Id by sheet name
    * @param xlsxPath The xlsx file path
    * @param sheetName The xlsx sheet name
    * @return
    */
  def lookupSheetIdByName(xlsxPath: String, sheetName: String): Option[String] = {

    val zis = extractStream(xlsxPath, workbook)
    val handler = new BaseHandler("sheet", "sheetId", 0)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }

    if(handler.getResult.nonEmpty) Option(handler.getResult.head) else None
  }

  private def buildInnerZipSheetFilePath(sheetsFolder: String, sheetId: String): String = {
    sheetsFolder + "/sheet" + sheetId + ".xml";
  }

  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = MAX_VALUE): ListBuffer[SSRawCell] = {

    //TODO serve una validazione dei parametri inseriti

    val sheetId = lookupSheetIdByName(xlsxPath, sheet)

    sheetId match {

      case some =>

        val sheetFileName = buildInnerZipSheetFilePath(sheets_folder, some.get)
        println(s"Reading sheet file at path $sheetFileName")

        val stylesList = lookupCellsStyles(xlsxPath)
        val handler = new SheetHandler(fromRow, toRow, stylesList)

        val zis = extractStream(xlsxPath, sheetFileName)
        if (zis.isDefined) {
          parser.parse(zis.get, handler)
        }
        handler.getResult

      case _ =>
        println(s"No sheet available with name $sheetId")
        ListBuffer[SSRawCell]()
    }
  }

  def lookupSharedString(xlsxPath: String, ids: Set[Int] = Set[Int]()): ListBuffer[String] = {

    val zis = extractStream(xlsxPath, shared_strings)
    val handler = new SharedStringsHandler(ids)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  private def lookupCellsStyles(xlsxPath: String): ListBuffer[SSCellStyle] = {

    val zis = extractStream(xlsxPath, styles)
    val handler = new StyleHandler

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  /**
    * In questo metodo viene effettuata una lettura dal file di sharedstring per ogni stringa, c'è già l'approccio a lista
    * usalo!
    * @param xlsxPath The xlsx file path
    * @param rawCellsList The cell list of the raw values
    * @return
    */
  def lookupValues(xlsxPath: String, rawCellsList: ListBuffer[SSRawCell]): List[SSMCell] = {

    rawCellsList.map(rawCell => {

      if("s".equals(rawCell.ctype)){

        if(rawCell.style==null) {
          //Shared string value
          val stringValue = lookupSharedString(xlsxPath, Set(rawCell.value.toInt))
          SSStringCell(rawCell.xy, rawCell.row, rawCell.column, stringValue.head)
        } else {
          parseFormatValue(xlsxPath, rawCell)
        }
      } else if(rawCell.ctype==null) {

        if(rawCell.style==null) {
          //Integer value
          SSIntegerCell(rawCell.xy, rawCell.row, rawCell.column, toInt(rawCell.value).getOrElse(0))
        } else {
          //Double value
          parseFormatValue(xlsxPath, rawCell)
        }
      } else {
        //Default String value
        SSStringCell(rawCell.xy, rawCell.row, rawCell.column, "buuuu")
      }
    }).toList
  }

  private def parseFormatValue(xlsxPath: String, cell: SSRawCell): SSMCell = {

    val style: SSCellStyle = cell.style

    style.numFmtId match {

      case 1 =>
        //Number into shared string table
        val stringValue = lookupSharedString(xlsxPath, Set(cell.value.toInt))
        SSIntegerCell(cell.xy, cell.row, cell.column, toInt(stringValue.head).getOrElse(0))
      case 164 => formatCurrencyCellValue(cell)
      case _ =>
        println(s"UNKNOWN numFmtId ${style.numFmtId}")
        SSDoubleCell(cell.xy, cell.row, cell.column, toDouble(cell.value).getOrElse(0.0))

    }
  }

  private def formatCurrencyCellValue(cell: SSRawCell): SSMCell = {

    cell.style.formatCode match {

      case "\"$\"#,##0" =>
        //val doubleVal = BigDecimal(cell.value).setScale(2, BigDecimal.RoundingMode.HALF_UP).toDouble
        //new SSCurrencyCell(cell.xy, cell.row, cell.column, "$", doubleVal)
        SSCurrencyCell(cell.xy, cell.row, cell.column, "$", toDouble(cell.value).getOrElse(0.0))
      case _ => SSDoubleCell(cell.xy, cell.row, cell.column, toDouble(cell.value).getOrElse(0.0))
    }
  }
}

object DocumentSaxParser {

  sealed trait SSMCell {
    val rowCol: String
    val rowNum: Int
    val colNum: Int
  }

  case class SSStringCell(rowCol: String, rowNum: Int, colNum: Int, value: String) extends SSMCell
  case class SSDoubleCell(rowCol: String, rowNum: Int, colNum: Int, value: Double) extends SSMCell
  case class SSIntegerCell(rowCol: String, rowNum: Int, colNum: Int, value: Int) extends SSMCell
  case class SSDateCell(rowCol: String, rowNum: Int, colNum: Int, value: Date) extends SSMCell
  case class SSCurrencyCell(rowCol: String, rowNum: Int, colNum: Int, currency: String, value: Double) extends SSMCell
}