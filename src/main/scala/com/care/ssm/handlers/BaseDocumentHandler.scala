package com.care.ssm.handlers

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

/**
  * Ci sono due possibilità dato un tag da ricercare:
  * - si cerca un attributo
  * - si cerca il contenuto del tag
  * @param targetTag
  * @param targetAttribute
  */
class BaseDocumentHandler(targetTag: String, targetAttribute: String = "") extends DefaultHandler{

  var result = ""
  var targetTagStarted = false
  var targetTagEnded = false
  var wordDone = false

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(wordDone) return

    if(qName.equals(targetTag)){
      //Starting event of target tag
      println(s"Start event of element: $qName")

      targetTagStarted = true

      if(isAttributeRequired) {
        //Attribute extraction
        result = attributes.getValue(targetAttribute)
        wordDone = true
      } else {
        //the result is the text within the tag
      }
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (wordDone) return

    if (qName.equals(targetTag)) {
      //Closing event of target tag
      println(s"End event of element: $qName")
      wordDone = true
    }
  }


  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

      if(!isAttributeRequired && targetTagStarted && !targetTagEnded) {
        result = new String(ch, start, length)
      }
  }

  private def isAttributeRequired: Boolean = {
    targetAttribute!=null && !targetAttribute.trim.isEmpty
  }

  def getResult(): String ={
    result
  }
}