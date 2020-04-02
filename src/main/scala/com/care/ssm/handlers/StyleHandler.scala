package com.care.ssm.handlers

import com.care.ssm.SSMUtils._
import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer

/**
  *
  * @author Massimo Caresana
  *
  * Handler for Shared Strings File
  */
class StyleHandler extends DefaultHandler{

  var result: ListBuffer[SSCellStyle] =  ListBuffer[SSCellStyle]()

  //Target Tags
  val parentTag = "cellXfs"
  var parentTagStarted = false
  var parentTagEnded = false
  val targetTag = "xf"
  var targetTagEnded = false

  //Target attributes
  val applyNumberFormatTag = "applyNumberFormat"
  val numFmtIdTag = "numFmtId"

  //Number formats Tags
  val numFmtsTag = "numFmts"
  val numFmtTag = "numFmt"
  var numFmtsEnded = false
  val formatCodeTag = "formatCode"

  //Number formats List (defined within <numFmts>)
  var numFormatsList = new ListBuffer[SSCellStyle]()

  //Position/Style Index
  var index = 0

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(parentTag.equals(qName)) {
      parentTagStarted = true
    }

    if(!numFmtsEnded && numFmtTag.equals(qName)) {
      //Starting numFmt tag
      val numFmtId = toInt(attributes.getValue(numFmtIdTag)).getOrElse(-1)
      val formatCode = attributes.getValue(formatCodeTag)
      //List is filled before the parsing of the style tags
      numFormatsList += SSCellStyle(numFmtId,formatCode)
    }

    if(parentTagStarted && targetTag.equals(qName)) {
      //Starting target tag
      val applyNumberFormatTh = attributes.getValue(applyNumberFormatTag)

      if (applyNumberFormatTh==null || applyNumberFormatTh.trim.isEmpty) {
        result += null
      } else {
        val numFormatId = toInt(attributes.getValue(numFmtIdTag)).getOrElse(-1)
        val formatCode = numFormatsList(Integer.valueOf(applyNumberFormatTh) - 1).formatCode
        result += SSCellStyle(numFormatId, formatCode)
      }
      index += 1
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if(!numFmtsEnded && numFmtsTag.equals(qName)){
      numFmtsEnded = true
    }

    if (!parentTagEnded && parentTag.equals(qName)){
      parentTagEnded = true
    }
  }

  private def workDone: Boolean = {
    parentTagEnded
  }

  def getResult: ListBuffer[SSCellStyle] = {
    result
  }
}

object StyleHandler {
  case class SSCellStyle(numFmtId: Int, formatCode: String)
}