package com.care.ssm

import java.util
import java.util.zip.ZipInputStream

import com.care.ssm.handlers.SheetHandler.SSCell
import com.care.ssm.handlers.{BaseDocumentHandler, BaseHandler, SharedStringsHandler, SheetHandler}
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
    } else {
      return result.get(0)
    }
  }

  def bubu(xlsxPath: String): util.ArrayList[String] ={
    extractData(xlsxPath, SSMUtils.styles, "cellStyle", "builtinId")
  }

  def readSheet(xlsxPath: String, sheet: String): util.ArrayList[SSCell] = {

    val sheetId = lookupSheetIdByName(xlsxPath, sheet)
    val sheetFileName= SSMUtils.sheets_folder + "/sheet" + sheetId + ".xml"

    println(s"Reading sheet file at path $sheetFileName")

    val zis = SSMUtils.extractStream(xlsxPath, sheetFileName)
    val handler = new SheetHandler()

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


  private def extractData(xlsxPath: String, filePath: String, tag: String, attribute: String = "", occurrence: Int = -1): util.ArrayList[String] = {

    val zis = SSMUtils.extractStream(xlsxPath, filePath)
    val handler = new BaseDocumentHandler(tag, attribute, occurrence)

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }

  def read(zis : ZipInputStream): Unit ={
    val handler = new BaseHandler
    parser.parse(zis,handler)
  }
}