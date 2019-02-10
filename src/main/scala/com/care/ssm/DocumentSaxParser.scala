package com.care.ssm

import java.util

import com.care.ssm.handlers.SheetHandler.SSCell
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import com.care.ssm.handlers.{BaseDocumentHandler, SharedStringsHandler, SheetHandler, StyleHandler}
import javax.xml.parsers.SAXParserFactory


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
      return Option(handler.getResult().get(0))
    } else {
      return None
    }
  }

  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = Integer.MAX_VALUE): util.ArrayList[SSCell] = {

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
        new util.ArrayList[SSCell]()
      }
    }
  }

  def lookupSharedString(xlsxPath: String, ids: Set[Int] = Set[Int]()): util.ArrayList[String] ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.shared_strings)
    val handler = new SharedStringsHandler(ids)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def lookupCellsStyles(xlsxPath: String): util.ArrayList[SSCellStyle] ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.styles)
    val handler = new StyleHandler

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }
}