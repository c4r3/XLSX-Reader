package com.c4r3.xlsx.reader

import java.lang.Integer.MAX_VALUE

import XLSXUtils._
import com.c4r3.xlsx.reader.handlers.{BaseHandler, SheetHandler, StyleHandler}
import com.c4r3.xlsx.reader.handlers.SheetHandler.Row
import com.c4r3.xlsx.reader.handlers.StyleHandler.CellStyle
import javax.xml.parsers.SAXParserFactory._
import javax.xml.parsers.{SAXParser, SAXParserFactory}
import org.slf4j
import org.slf4j.LoggerFactory


/**
  * @author C4r3
  */
class XLSXParser {

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

  def readSheet(xlsxPath: String, sheet: String, fromRow: Int = 0, toRow: Int = MAX_VALUE): List[Row] = {

    validateRangeBounds(fromRow, toRow)

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
  }

  private def validateRangeBounds(fromRow: Int, toRow: Int): Unit = {

    if(fromRow<0) throw new IllegalArgumentException("Range lower bound can't be negative")
    if(toRow<0) throw new IllegalArgumentException("Range upper bound can't be negative")
    if(toRow<fromRow) throw new IllegalArgumentException("Range upper bound can't be lower than lower one")

    if(fromRow!=0 && toRow != Integer.MAX_VALUE) {
      logger.debug("Start reading from row {} to {}", fromRow, toRow - 1)
    }
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