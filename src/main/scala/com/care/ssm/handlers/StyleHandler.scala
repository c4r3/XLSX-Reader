package com.care.ssm.handlers

import com.care.ssm.SSMUtils._
import com.care.ssm.handlers.StyleHandler.CellStyle
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

  val result: ListBuffer[CellStyle] =  ListBuffer[CellStyle]()

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
  var numFormatsList = new ListBuffer[CellStyle]()

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
      numFormatsList += CellStyle(numFmtId,formatCode)
    }

    if(parentTagStarted && targetTag.equals(qName)) {

      toInt(attributes.getValue(numFmtIdTag)) match {
        case Some(num) =>

          //Se il numFmtId è presente nella numFormatList allora prendo quel valore, altrimenti significa che è uno stile standard (definito a priori)
          numFormatsList.find( sCell => sCell.numFmtId == num) match {

            case Some(style) => result += style
            case None =>
              val currentNumFormat = lookupFormatCode(num)
              result += CellStyle(num, currentNumFormat)
          }
        case None => result += null
      }

      index += 1
    }
  }

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
        print(s"UNKNOWN numfmtId $c")
        null
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
  case class CellStyle(numFmtId: Int, formatCode: String)
}