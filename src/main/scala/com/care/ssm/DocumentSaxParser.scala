package com.care.ssm

import java.util
import java.util.zip.ZipInputStream

import com.care.ssm.handlers.{BaseDocumentHandler, BaseHandler}
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
  def lookupSheetIdByName(xlsxPath: String, sheetName: String): util.ArrayList[String] ={
    extractData(xlsxPath, SSMUtils.workbook, "sheet","sheetId")
  }

  def bubu(xlsxPath: String): util.ArrayList[String] ={
    extractData(xlsxPath, SSMUtils.styles, "cellStyle", "builtinId")
  }



  private def extractData(xlsxPath: String, filePath: String, tag: String, attribute: String = ""): util.ArrayList[String] = {

    val zis = SSMUtils.extractStream(xlsxPath, filePath)
    val handler = new BaseDocumentHandler(tag, attribute)

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