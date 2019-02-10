package com.care.ssm.handlers

import java.util

import com.care.ssm.handlers.StyleHandler.{SSCellNumberFormat, SSCellStyle}
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer

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
  val numFmtIdTag = "numFmtId"

  //Number formats Tags
  val numFmtsTag = "numFmts"
  val numFmtTag = "numFmt"
  var numFmtsEnded = false
  val formatCodeTag = "formatCode"

  //Number formats List (defined within <numFmts>)
  var numFormatsList = new ListBuffer[SSCellNumberFormat]()

  //Position/Style Index
  var index = 0

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(parentTag.equals(qName)) {
      parentTagStarted = true
    }

    if(!numFmtsEnded && numFmtTag.equals(qName)) {
      //Starting numFmt tag
      val numFmtId = attributes.getValue(numFmtIdTag)
      val formatCode = attributes.getValue(formatCodeTag)
      //List is filled before the parsing of the style tags
      numFormatsList += new SSCellNumberFormat(numFmtId,formatCode)
    }

    if(parentTagStarted && targetTag.equals(qName)) {
      //Starting target tag
      val applyFormatTh = attributes.getValue(applyNumberFormatTag)
      val numFormatId = attributes.getValue(numFmtIdTag)
      val formatCode = if (applyFormatTh==null || applyFormatTh.trim.isEmpty) {
        "-1"
      } else {
         numFormatsList(Integer.valueOf(applyFormatTh) - 1).formatCode
      }

      result.add(new SSCellStyle(index, applyFormatTh, numFormatId, formatCode))
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
  case class SSCellStyle(index: Int, applyNumberFormat: String, numFmtId: String, formatCode: String)

  case class SSCellNumberFormat(numFmtId: String, formatCode: String)
}