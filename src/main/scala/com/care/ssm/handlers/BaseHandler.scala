package com.care.ssm.handlers

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

class BaseHandler(targetTag: String = "") extends DefaultHandler{

  var buffer = ""
  var wordDone = false

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(wordDone) return

    if(qName.equals(targetTag)){
      //Starting event of target tag
      println(s"Start: $qName")

      //Attribute extraction
      buffer = attributes.getValue("sheetId")
      wordDone = true
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if(wordDone) return

    //Closing event of target tag
    if(qName.equals(targetTag) ) println(s"End: $qName")
  }


  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

     //buffer = new String(ch, start, length)
  }

  def getResult(): String ={
    buffer
  }
}