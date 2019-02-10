package com.care.ssm.handlers

import java.util

import com.care.ssm.SSMUtils
import com.care.ssm.handlers.SheetHandler.SSCell
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.xml.Null

/**
  * Handler for Shared Strings File
  * @param indexes
  */
class SheetHandler(fromRow: Int = 0, toRow: Int = Integer.MAX_VALUE, stylesList: util.ArrayList[SSCellStyle] = new util.ArrayList[SSCellStyle]()) extends DefaultHandler{

  var result = new util.ArrayList[SSCell]

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
      cellRowNum = Integer.valueOf(rawRowNumValue)
    }

    if(isNotRequiredRow) return

    if(valueTag.equals(qName)){
      valueTagStarted = true
    }

    if(targetTag.equals(qName)) {
      //Starting "c" tag, extraction of the attributes
      cellXY = attributes.getValue(xyAttr)
      cellStyle = toInt(attributes.getValue(styleAttr)).getOrElse(-1)
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
      val style: SSCellStyle = getStyle(stylesList, cellStyle).getOrElse(null) //se non c'è style si va in lookup nella shared string o è un numerico
      result.add(new SSCell(cellRowNum, SSMUtils.calculateColumn(cellXY, cellRowNum), cellXY, cellType, new String(ch, start, length), style))

      valueTagStarted = false
      cellXY = ""
      cellStyle = 0
      cellType = ""
    }
  }

  def toInt(s: String): Option[Int] = {
    try {
      Some(s.toInt)
    } catch {
      case e: Exception => None
    }
  }

  def getStyle(stylesList: util.ArrayList[SSCellStyle], styleIndex: Int): Option[SSCellStyle] = {

    if (stylesList != null && stylesList.size > styleIndex && styleIndex >= 0) {
      Some(stylesList.get(styleIndex))
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

  def getResult: util.ArrayList[SSCell] ={
    result
  }
}

object SheetHandler {

  case class SSCell(row: Int, column: Int, xy: String, ctype: String, value: String, style: SSCellStyle)
}