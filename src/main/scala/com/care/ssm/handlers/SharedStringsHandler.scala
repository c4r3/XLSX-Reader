package com.care.ssm.handlers

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer

/**
  *
  * @author Massimo Caresana
  *
  *         Handler for Shared Strings File
  * @param indexes the required indexes of the strings
  */
class SharedStringsHandler(indexes: Set[Int] = Set[Int]()) extends DefaultHandler {

  var result: ListBuffer[String] = ListBuffer[String]()
  val targetTag = "t"
  var counter: Int = 0
  var targetTagEnded: Boolean = false

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if (workDone) return

    if (isTTag(qName)) {
      targetTagEnded = false
    }
  }

  private def isTTag(tag: String): Boolean = {
    targetTag.equals(tag);
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if (isTTag(qName)) {
      counter += 1
      targetTagEnded = true
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone) return

    if (indexes.isEmpty || indexes.contains(counter)) {
      result += new String(ch, start, length)
    }
  }

  private def workDone: Boolean = {
    indexes.nonEmpty && result.length == indexes.size
  }

  def getResult: ListBuffer[String] = {
    result
  }
}