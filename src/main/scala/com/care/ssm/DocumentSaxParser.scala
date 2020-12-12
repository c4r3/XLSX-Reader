package com.care.ssm

import java.lang.Integer.MAX_VALUE
import java.sql.Date

import com.care.ssm.DocumentSaxParser._
import com.care.ssm.SSMUtils.SSCellType.{SSCellType => _}
import com.care.ssm.SSMUtils.{SSCellType, _}
import com.care.ssm.handlers.SheetHandler.{Cell, Row}
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import com.care.ssm.handlers.{BaseHandler, SharedStringsHandler, SheetHandler, StyleHandler}
import javax.xml.parsers.{SAXParser, SAXParserFactory}

import scala.collection.mutable.ListBuffer


/**
  * @author Massimo Caresana
  */
class DocumentSaxParser {

  val factory: SAXParserFactory = SAXParserFactory.newInstance
  val parser: SAXParser = factory.newSAXParser

  /**
    * Look up sheet Id by sheet name
    *
    * @param xlsxPath  The xlsx file path
    * @param sheetName The xlsx sheet name
    * @return
    */
  def lookupSheetIdByName(xlsxPath: String, sheetName: String): Option[String] = {

    val zis = extractStream(xlsxPath, workbook)
    val handler = new BaseHandler("sheet", "sheetId", 0)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }

    if (handler.getResult.nonEmpty) Option(handler.getResult.head) else None
  }

  private def buildInnerZipSheetFilePath(sheetsFolder: String, sheetId: String): String =
    sheetsFolder + "/sheet" + sheetId + ".xml"

  //TODO serve una validazione dei parametri inseriti
  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = MAX_VALUE): List[Row] = {

    val sheetId = lookupSheetIdByName(xlsxPath, sheet)
    sheetId match {

      case some =>

        val sheetFileName = buildInnerZipSheetFilePath(sheets_folder, some.get)
        println(s"Reading sheet file at path $sheetFileName")

        //Style data usually is a small amount of data, so it can be read once
        val stylesList = lookupCellsStyles(xlsxPath)
        val handler = new SheetHandler(fromRow, toRow, stylesList, xlsxPath)

        val zis = extractStream(xlsxPath, sheetFileName)
        if (zis.isDefined) {
          parser.parse(zis.get, handler)
        }
        handler.getResult

      case _ =>
        println(s"No sheet available with name $sheetId")
        List[Row]()
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

  private def lookupCellsStyles(xlsxPath: String): List[SSCellStyle] = {

    val zis = extractStream(xlsxPath, styles)
    val handler = new StyleHandler

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

//  /**
//    * In questo metodo viene effettuata una lettura dal file di sharedstring per ogni stringa, c'è già l'approccio a lista
//    * usalo!
//    *
//    * @param xlsxPath     The xlsx file path
//    * @param rawCellsList The cell list of the raw values
//    * @return
//    */
//  def parseRawCells(xlsxPath: String, rawCellsList: List[SSRawCell]): List[SSMCell] =
//    rawCellsList.flatMap(rawCell => parseRawCell(xlsxPath, rawCell))


//  def parseRawCell(xlsxPath: String, rawCell: SSRawCell): Option[SSMCell] = {
//
//    detectCellType(rawCell.ctype) match {
//
//      case SSCellType.InlineString => Some(parseInlineStringCell(rawCell))
//      case SSCellType.Date => None //TODO aggiungere la casistica
//      case SSCellType.Double => Some(createSSDoubleCell(rawCell, rawCell.value))
//      case SSCellType.Error => None //TODO aggiungere
//      case SSCellType.Number => Some(createSSLongCellValue(rawCell))
//
//      case SSCellType.SharedString | _ =>
//        //TODO se è di tipo SharedString vuol dire che ha il valore nella SharedStringsTable
//        val sharedStringValues = lookupSharedString(xlsxPath, Set(rawCell.value.toInt))
//
//        if (noStylePresent(rawCell.style)) {
//          Some(createSSStringCellValue(rawCell, sharedStringValues.head))
//        } else {
//          parseFormatValue(rawCell, sharedStringValues.head)
//        }
//      case _ => {
//        System.err.println(s"No type detected for raw cell $rawCell")
//        None
//      }
//    }
//  }
//
//  //Inline String value
//  def parseInlineStringCell(rawCell: SSRawCell): SSMCell =
//    SSStringCell(rawCell.xy, rawCell.row, rawCell.column, rawCell.value)
//
//  def noStylePresent(style: SSCellStyle): Boolean = style == null //TODO introdurre un None
//
//  //reference http://www.ecma-international.org/publications/standards/Ecma-376.htm
//  //pag 2127, pdf 1st ed part 4
//  private def parseFormatValue(cell: SSRawCell, stringValue: String): Option[SSMCell] = {
//
//    if(cell.style == null) {
//      Some(createSSStringCellValue(cell, stringValue))
//    } else {
//
//      if(cell.style.formatCode == "0.00%") {
//        Some(createSSDoubleCell(cell, stringValue))
//      } else if(cell.style.formatCode == "#,##0.00 \"€\"") {
//        //FIXME deve essere gestito con il caso di valore dalla shared string e anche con la valuta opportuna
//        //Some(createSSCurrencyCellValue(cell, stringValue))
//        Some(createSSCurrencyCellValue(cell))
//      } else if(cell.style.formatCode == "#,##0.00\\ \"€\"") {
//        //FIXME deve essere gestito con il caso di valore dalla shared string e anche con la valuta opportuna
//        //Some(createSSCurrencyCellValue(cell, stringValue))
//        Some(createSSCurrencyCellValue(cell))
//      } else if(cell.style.formatCode == "@") {
//        Some(createSSStringCellValue(cell, stringValue))
//      } else if(cell.style.formatCode == "0.00E+00") {
//        Some(createSSDoubleCell(cell, stringValue))
//      } else if(cell.style.formatCode == "#,##0.00") {
//        Some(createSSDoubleCell(cell, stringValue))
//      } else if(cell.style.formatCode == "0.00%") {
//        Some(createSSDoubleCell(cell, stringValue))
//      } else if(cell.style.formatCode == "0.0000") {
//        Some(createSSDoubleCell(cell, stringValue))
//      } else {
//        None
//      }
//    }
//
//        //TODO da controllare i cell.value, probabilmente vanno presi dalla sharedStrings
////      case 0 => createSSDoubleCell(cell, stringValue) //General
////      case 1 => createSSIntegerCell(cell) //0
////      case 2 => createSSDoubleCell(cell, stringValue) //0.00
////      case 3 => createSSDoubleCell(cell, stringValue) //#,##0
////      case 4 => createSSDoubleCell(cell, stringValue) //#,##0.00
////      case 9 => createSSDoubleCell(cell, cell.value) //0%
////      case 10 => createSSDoubleCell(cell, cell.value) //0.00%
////      case 11 => createSSDoubleCell(cell, cell.value) //0.00E+00
////      case 12 => null //# ?/?
////      case 13 => null //# ??/??
////      case 14 => null //mm-dd-yy
////      case 15 => null //d-mmm-yy
////      case 16 => null //d-mmm
////      case 17 => null //mmm-yy
////      case 18 => null //h:mm AM/PM
////      case 19 => null //h:mm:ss AM/PM
////      case 20 => null //h:mm
////      case 21 => null //h:mm:ss
////      case 22 => null //m/d/yy h:mm
////      case 37 => createSSDoubleCell(cell, cell.value) //#,##0 ;(#,##0)
////      case 38 => createSSDoubleCell(cell, cell.value) //#,##0 ;[Red](#,##0)
////      case 39 => createSSDoubleCell(cell, cell.value) //#,##0.00;(#,##0.00)
////      case 40 => createSSDoubleCell(cell, cell.value) //#,##0.00;[Red](#,##0.00)
////      case 45 => null //mm:ss
////      case 46 => null //[h]:mm:ss
////      case 47 => null //mmss.0
////      case 48 => createSSDoubleCell(cell, cell.value) //##0.0E+0
////      case 49 => null //@
////        //TODO da qui in giù potrebbero essere da controllare all'interno del tag "numFmt" in quanto possono essere custom
////      case 164 => formatCurrencyCellValue(cell)
////      case 165 => null //forse integer o float https://github.com/jsonkenl/xlsxir/issues/17
////      case 166 => formatCurrencyCellValue(cell) //$#,##0.00
////      case 167 => null //forse date https://github.com/jsonkenl/xlsxir/issues/17
////      case _ =>
////        println(s"UNKNOWN numFmtId ${cell.style.numFmtId}")
////        createSSDoubleCell(cell, cell.value)
////    }
////
////    Some(res)
//  }
//
//  //TODO occorre una completa gestione delle valute
//  private def formatCurrencyCellValue(cell: SSRawCell): SSMCell = {
//
//    cell.style.formatCode match {
//      case "#,##0.00\\ &quot;€&quot;" | "\"$\"#,##0" => createSSCurrencyCellValue(cell)
//      case _ => createSSDoubleCell(cell, cell.value)
//    }
//  }

  //EVALUATED CELLS VALUES AND CREATION
//  def createSSDoubleCell(cell: SSRawCell, stringValue: String): SSMCell = {
//
//    SSDoubleCell(cell.xy, cell.row, cell.column, toDouble(stringValue).getOrElse(0.0))
//  }
//
//  def createSSIntegerCell(cell: SSRawCell): SSMCell = {
//    SSIntegerCell(cell.xy, cell.row, cell.column, toInt(cell.value).getOrElse(0))
//  }
//
//  //TODO correggere con più valute
//  def createSSCurrencyCellValue(cell: SSRawCell): SSMCell = {
//    SSCurrencyCell(cell.xy, cell.row, cell.column, "$", toDouble(cell.value).getOrElse(0.0))
//  }
//
//  def createSSStringCellValue(cell: SSRawCell, stringValue: String): SSMCell = {
//
//    if(stringValue.isEmpty) { //No value found --> Integer Cell
//      createSSIntegerCell(cell)
//    } else {
//      SSStringCell(cell.xy, cell.row, cell.column, stringValue)
//    }
//  }
//
//  def createSSLongCellValue(cell: SSRawCell): SSMCell = {
//    SSLongCell(cell.xy, cell.row, cell.column, toLong(cell.value).getOrElse(0))
//  }
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

  case class SSLongCell(rowCol: String, rowNum: Int, colNum: Int, value: Long) extends SSMCell
}