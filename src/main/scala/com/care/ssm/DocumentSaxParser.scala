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
  def lookupSheetIdByName(xlsxPath: String, sheetName: String): String ={
    val result = extractData(xlsxPath, SSMUtils.workbook, "sheet","sheetId", 0)
    if(result.isEmpty) {
      return ""
    } else { //TODO correggere sta cazzata, deve avere un solo ritorno
      return result.get(0)
    }
  }

  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = Integer.MAX_VALUE): util.ArrayList[SSCell] = {

    val sheetId = lookupSheetIdByName(xlsxPath, sheet)
    val sheetFileName = SSMUtils.sheets_folder + "/sheet" + sheetId + ".xml"

    println(s"Reading sheet file at path $sheetFileName")

    val zis = SSMUtils.extractStream(xlsxPath, sheetFileName)
    val handler = new SheetHandler(fromRow, toRow)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
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

  //TODO renderlo un unico metodo passandogli l'handler da fuori (deve essere definita una trait con il metodo getResult per gli effetti di bordo)
  private def extractData(xlsxPath: String, filePath: String, tag: String, attribute: String = "", occurrence: Int = -1): util.ArrayList[String] = {

    val zis = SSMUtils.extractStream(xlsxPath, filePath)
    val handler = new BaseDocumentHandler(tag, attribute, occurrence)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }
}