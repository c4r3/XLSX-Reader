package com.c4r3.ssm.handlers

import SheetHandler.toInt
import StyleHandler.{CellStyle, lookupFormatCode}
import org.slf4j
import org.slf4j.LoggerFactory
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

  val logger: slf4j.Logger = LoggerFactory.getLogger(this.getClass)

  val result: ListBuffer[CellStyle] =  ListBuffer[CellStyle]()

  //Target Tags
  val parentTag = "cellXfs"
  val targetTag = "xf"

  var parentTagStarted = false
  var parentTagEnded = false
  var targetTagEnded = false

  //Target attributes
  val applyNumberFormatTag = "applyNumberFormat"
  val numFmtIdTag = "numFmtId"

  //Number formats Tags
  val numFmtsTag = "numFmts"
  val numFmtTag = "numFmt"
  val formatCodeTag = "formatCode"

  var numFmtsEnded = false

  //Number formats List (defined within <numFmts>)
  var numFormatsList = new ListBuffer[CellStyle]()

  //Position/Style Index
  var index = 0

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(parentTag.equals(qName)) {
      parentTagStarted = true
    }

    if(!numFmtsEnded && numFmtTag.equals(qName)) {

      val numFmtId = toInt(attributes.getValue(numFmtIdTag)).getOrElse(-1)
      val formatCode = attributes.getValue(formatCodeTag)

      numFormatsList += CellStyle(numFmtId,formatCode, isFormatNumberApply = false)
    }

    if(parentTagStarted && targetTag.equals(qName)) {

      val isFormatNumberApply = isApplyNumberFormatRequired(attributes.getValue(applyNumberFormatTag))

      toInt(attributes.getValue(numFmtIdTag)) match {
        case Right(num) =>

          numFormatsList.find(sCell => sCell.numFmtId == num) match {

            case Some(style) => result += CellStyle(style.numFmtId, style.formatCode, isFormatNumberApply)
            case None =>
              val currentNumFormat = lookupFormatCode(num)
              result += CellStyle(num, currentNumFormat, isFormatNumberApply)
          }
        case Left(_) => result += null
      }

      index += 1
    }
  }

  def isApplyNumberFormatRequired(applyString: String): Boolean = {

    if(applyString==null || applyString.trim.isEmpty) return false

    toInt(applyString) match {
      case Right(applyVal) => applyVal==1
      case Left(_) => false
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

  private def workDone: Boolean = parentTagEnded

  def getResult: List[CellStyle] = result.toList
}

object StyleHandler {

  val logger: slf4j.Logger = LoggerFactory.getLogger(StyleHandler.getClass)

  case class CellStyle(numFmtId: Int, formatCode: String, isFormatNumberApply: Boolean)

  def lookupFormatCode(c : Int): String =
    c match {
      case 0 => "General"
      case 1 => "0"
      case 2 => "0.00"
      case 3 => "#,##0"
      case 4 => "#,##0.00"
      case 9 => "0%"
      case 10 => "0.00%"
      case 11 => "0.00E+00"
      case 12 => "#?/?"
      case 13 => "#??/??"
      case 14 => "mm-dd-yy"
      case 15 => "d-mmm-yy"
      case 16 => "d-mmm"
      case 17 => "mmm-yy"
      case 18 => "h:mm AM/PM"
      case 19 => "h:mm:ss AM/PM"
      case 20 => "h:mm"
      case 21 => "h:mm:ss"
      case 22 => "m/d/yy h:mm"
      case 37 => "#,##0;(#,##0)"
      case 38 => "#,##0 ;[Red](#,##0)"
      case 39 => "#,##0.00;(#,##0.00)"
      case 40 => "#,##0.00;[Red](#,##0.00)"
      case 45 => "mm:ss"
      case 46 => "[h]:mm:ss"
      case 47 => "mmss.0"
      case 48 => "##0.0E+0"
      case 49 => "@"
      case _ =>
        logger.error("Unknown numfmtId {}", c)
        null
    }
}