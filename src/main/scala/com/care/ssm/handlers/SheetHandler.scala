package com.care.ssm.handlers

import java.lang.Integer._

import com.care.ssm.SSMUtils
import com.care.ssm.SSMUtils.calculateColumn
import com.care.ssm.handlers.SheetHandler.SSRawCell
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer


/**
  * Handler for Shared Strings File
  * @param fromRow
  * @param toRow
  * @param stylesList
  */
class SheetHandler(fromRow: Int = 0, toRow: Int = MAX_VALUE, stylesList: ListBuffer[SSCellStyle] = ListBuffer[SSCellStyle]()) extends DefaultHandler{

  var result = ListBuffer[SSRawCell]()

  //Target and value tags
  val targetTag = "c"
  val valueTag = "v"
  val rowTag = "row"

  //Attributes tags
  val rowNumAttr = "r"
  val xyAttr = "r"
  val styleAttr = "s"
  val typeAttr = "t"

  //Flags
  var cellEnded = false
  var valueTagStarted = false

  //Temporary variables
  var cellRowNum = 0
  var cellXY = ""
  var cellStyle = 0
  var cellType = ""

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(rowTag.equals(qName)) {
      val rawRowNumValue = attributes.getValue(rowNumAttr)
      cellRowNum = valueOf(rawRowNumValue)
    }

    if(isNotRequiredRow) return

    if(valueTag.equals(qName)){
      valueTagStarted = true
    }

    if(targetTag.equals(qName)) {
      //Starting "c" tag, extraction of the attributes
      cellXY = attributes.getValue(xyAttr)
      cellStyle = SSMUtils.toInt(attributes.getValue(styleAttr)).getOrElse(-1)
      cellType = attributes.getValue(typeAttr)
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {
    //if (workDone || isNotRequiredRow) return
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone || isNotRequiredRow) return

    if(valueTagStarted) {

      //Flushing buffer & reset temporary stuff
      val style: SSCellStyle = getStyle(stylesList, cellStyle).orNull //se non c'è style si va in lookup nella shared string o è un numerico
      result+= SSRawCell(cellRowNum, calculateColumn(cellXY, cellRowNum), cellXY, cellType, new String(ch, start, length), style)

      valueTagStarted = false
      cellXY = ""
      cellStyle = 0
      cellType = ""
    }
  }

  def getStyle(stylesList: ListBuffer[SSCellStyle], styleIndex: Int): Option[SSCellStyle] = {

    if (stylesList != null && stylesList.size > styleIndex && styleIndex >= 0) {
      Some(stylesList(styleIndex))
    } else {
      None
    }
  }

  private def isNotRequiredRow: Boolean = {
    !(fromRow.toInt <= cellRowNum && cellRowNum <= toRow.toInt)
  }

  private def workDone: Boolean = {
    cellRowNum > toRow.toInt
  }

  def getResult: ListBuffer[SSRawCell] ={
    result
  }
}

object SheetHandler {

  case class SSRawCell(row: Int, column: Int, xy: String, ctype: String, value: String, style: SSCellStyle)
}