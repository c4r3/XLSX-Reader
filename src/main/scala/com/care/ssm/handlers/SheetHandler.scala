package com.care.ssm.handlers

import java.util

import com.care.ssm.handlers.SheetHandler.SSCell
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

/**
  * Handler for Shared Strings File
  * @param indexes
  */
class SheetHandler(fromRow: Int = 0, toRow: Int = Integer.MAX_VALUE) extends DefaultHandler{

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
  var cellStyle = ""
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
      cellStyle = attributes.getValue(styleAttr)
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
      result.add(new SSCell(cellRowNum, cellXY, cellStyle, cellType, new String(ch, start, length)))

      valueTagStarted = false
      cellXY = ""
      cellStyle = ""
      cellType = ""
    }
  }

  private def isNotRequiredRow: Boolean = {
    !(fromRow.toInt <= cellRowNum && cellRowNum <= toRow.toInt)
  }


  private def workDone: Boolean = {
    false //TODO da completare
  }

  def getResult: util.ArrayList[SSCell] ={
    result
  }
}

object SheetHandler {

  //TODO devono essere specializzate le classi con i vari tipi in overloading,verrà fatto dopo che si capisce
  //TODO la logica per determinare l'entità dagli attributi
  case class SSCell(rowNum: Int, xy: String, style: String, ctype: String, value: String)
}