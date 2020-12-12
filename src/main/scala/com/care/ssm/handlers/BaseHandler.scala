package com.care.ssm.handlers

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer

/**
  *
  * @author Massimo Caresana
  *
  * Ci sono due possibilità dato un tag da ricercare:
  * - si cerca un attributo
  * - si cerca il contenuto del tag
  * @param targetTag The target tag
  * @param targetAttribute The target tag's attribute
  * @param occurrence <0 all values
  */
class BaseHandler(targetTag: String, targetAttribute: String = "", occurrence: Int = -1) extends DefaultHandler {

  var result: ListBuffer[String] = ListBuffer[String]()
  var targetTagStarted = false
  var targetTagEnded = false
  var wordDone = false

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if (workDone) return

    if (qName.equals(targetTag)) {
      //Starting event of target tag
      println(s"Start event of element: $qName")

      targetTagStarted = true

      if (isAttributeRequired) {
        //Attribute extraction
        result += attributes.getValue(targetAttribute)
      } else {
        //the result is the text within the tag
      }
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if (qName.equals(targetTag)) {
      //Closing event of target tag
      println(s"End event of element: $qName")
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (!isAttributeRequired && targetTagStarted && !targetTagEnded) {
      result += new String(ch, start, length)
    }
  }

  //negative occurrence -> all values: [0,[ -> the corresponding value(s)
  private def workDone: Boolean = occurrence > 0 && result.size == occurrence

  private def isAttributeRequired: Boolean = targetAttribute != null && !targetAttribute.trim.isEmpty

  def getResult: List[String] = result.toList
}