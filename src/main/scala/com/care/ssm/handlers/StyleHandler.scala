package com.care.ssm.handlers

import java.util

import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

/**
  * Handler for Shared Strings File
  * @param indexes
  */
class StyleHandler extends DefaultHandler{

  var result =  new util.ArrayList[SSCellStyle]

  //Target Tags
  val parentTag = "cellXfs"
  var parentTagStarted = false
  var parentTagEnded = false
  val targetTag = "xf"
  var targetTagEnded = false

  //Target attributes
  val applyNumberFormatTag = "applyNumberFormat"
  val numFmtId = "numFmtId"

  var counter = 0

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(parentTag.equals(qName)) {
      parentTagStarted = true
    }

    if(parentTagStarted && targetTag.equals(qName)) {
      //Starting target tag
      val applyFormatTh: String = attributes.getValue(applyNumberFormatTag)
      val numFormatId = attributes.getValue(numFmtId)
      val formatCode = "buuuu"
      result.add(new SSCellStyle(counter, applyFormatTh, numFormatId, formatCode))
      counter += 1
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if (parentTag.equals(qName)){
      parentTagEnded = true
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone) return

  }

  private def workDone: Boolean = {
    parentTagEnded
  }

  def getResult: util.ArrayList[SSCellStyle] ={
    result
  }
}

object StyleHandler {
  //TODO da portare i valori string ad interi
  case class SSCellStyle(position: Int, numberFormat: String, numFmtId: String, formatCode: String)

}