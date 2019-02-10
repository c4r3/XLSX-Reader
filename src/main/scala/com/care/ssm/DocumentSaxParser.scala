package com.care.ssm

import java.sql.Date

import com.care.ssm.DocumentSaxParser.{SSMCell, SSStringCell}
import com.care.ssm.handlers.SheetHandler.SSRawCell
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import com.care.ssm.handlers.{BaseDocumentHandler, SharedStringsHandler, SheetHandler, StyleHandler}
import javax.xml.parsers.SAXParserFactory

import scala.collection.mutable.ListBuffer


class DocumentSaxParser {

  val factory =  SAXParserFactory.newInstance
  val parser = factory.newSAXParser

  /**
    * Look up sheet Id by sheet name
    * @param xlsxPath
    * @param sheetName
    * @return
    */
  def lookupSheetIdByName(xlsxPath: String, sheetName: String): Option[String] ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.workbook)
    val handler = new BaseDocumentHandler("sheet", "sheetId", 0)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    if(!handler.getResult.isEmpty) {
      return Option(handler.getResult()(0))
    } else {
      return None
    }
  }

  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = Integer.MAX_VALUE): ListBuffer[SSRawCell] = {

    val sheetId = lookupSheetIdByName(xlsxPath, sheet)
    sheetId match {

      case some => {
        val sheetFileName = SSMUtils.sheets_folder + "/sheet" + some.get + ".xml"

        println(s"Reading sheet file at path $sheetFileName")

        //Reading style
        val stylesList = lookupCellsStyles(xlsxPath)

        val zis = SSMUtils.extractStream(xlsxPath, sheetFileName)
        val handler = new SheetHandler(fromRow, toRow, stylesList)

        if (zis.isDefined) {
          parser.parse(zis.get, handler)
        }
        handler.getResult
      }
      case _ => {
        println(s"No sheet available with name $sheetId")
        ListBuffer[SSRawCell]()
      }
    }
  }

  def lookupSharedString(xlsxPath: String, ids: Set[Int] = Set[Int]()): ListBuffer[String] ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.shared_strings)
    val handler = new SharedStringsHandler(ids)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def lookupCellsStyles(xlsxPath: String): ListBuffer[SSCellStyle] ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.styles)
    val handler = new StyleHandler

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def lookupValues(xlsxPath: String, rawCellsList: ListBuffer[SSRawCell]): List[SSMCell] ={
    //TODO aggiungere una LRU Cache
    rawCellsList.map(rawCell => {

      if("s".equals(rawCell.ctype) && rawCell.style==null){
        //Shared string value
        val stringValue: ListBuffer[String] = lookupSharedString(xlsxPath, Set(rawCell.value.toInt))
        new SSStringCell(rawCell.xy, rawCell.row, rawCell.column, stringValue(0))
      } else {
        //FIXME da completare con tutte le casistiche desiderate
        new SSStringCell(rawCell.xy, rawCell.row, rawCell.column, "buuuu")
      }
    }).toList
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