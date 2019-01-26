package com.care.ssm

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

class SaxHandler extends DefaultHandler{

  var buffer = ""

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    println("Start: " + qName)
  }


  override def endElement(uri: String, localName: String, qName: String): Unit = {

    println("End: " + qName)
  }


  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    buffer = new String(ch, start, length)
    println(buffer)
  }
}