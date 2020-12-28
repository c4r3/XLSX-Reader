package com.care.ssm

import java.lang.Integer.MAX_VALUE

import com.care.ssm.SSMUtils._
import com.care.ssm.handlers.SheetHandler.Row
import com.care.ssm.handlers.StyleHandler.CellStyle
import com.care.ssm.handlers.{BaseHandler, SheetHandler, StyleHandler}
import javax.xml.parsers.SAXParserFactory._
import javax.xml.parsers.{SAXParser, SAXParserFactory}
import org.slf4j
import org.slf4j.LoggerFactory


/**
  * @author Massimo Caresana
  */
class DocumentSaxParser {

  val logger: slf4j.Logger = LoggerFactory.getLogger(this.getClass)

  val factory: SAXParserFactory = newInstance
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

    handler.getResult.headOption
  }

  private def innerZipSheetFilePath(sheetsFolder: String, sheetId: String): String =
    sheetsFolder + "/sheet" + sheetId + ".xml"

  //TODO serve una validazione dei parametri inseriti
  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = MAX_VALUE): List[Row] =

    lookupSheetIdByName(xlsxPath, sheet) match {

      case Some(id) =>

        val sheetFileName = innerZipSheetFilePath(sheets_folder, id)
        logger.debug("Reading sheet file at path {}", sheetFileName)

        //Style data usually is a small amount of data, so it can be read just once
        val stylesList = lookupCellsStyles(xlsxPath)
        val handler = new SheetHandler(fromRow, toRow, stylesList, xlsxPath)

        val zis = extractStream(xlsxPath, sheetFileName)
        if (zis.isDefined) {
          parser.parse(zis.get, handler)
        }
        handler.getResult

      case _ =>
        logger.warn("No sheet available with name {}", sheet)
        List[Row]()
    }

  private def lookupCellsStyles(xlsxPath: String): List[CellStyle] = {

    val zis = extractStream(xlsxPath, styles)
    val handler = new StyleHandler

    if (zis.isDefined) {
      parser.parse(zis.get, handler)
    }
    handler.getResult
  }
}