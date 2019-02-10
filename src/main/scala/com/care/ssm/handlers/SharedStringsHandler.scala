package com.care.ssm.handlers

import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

import scala.collection.mutable.ListBuffer

/**
  * Handler for Shared Strings File
  * @param indexes
  */
class SharedStringsHandler(indexes: Set[Int] = Set[Int]()) extends DefaultHandler{

  var result =  ListBuffer[String]()
  val targetTag = "t"
  var counter: Int = 0
  var targetTagEnded: Boolean = false

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(targetTag.equals(qName)) {
      //Starting target tag
      targetTagEnded = false
      //println(s"Starting target tag: $targetTag")
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if(targetTag.equals(qName)) {
      //Closing target tag
      counter += 1
      targetTagEnded = true
      //println(s"Ending target tag: $targetTag")
    }
  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone) return

    if(indexes.isEmpty || indexes.contains(counter)) {
      result += new String(ch, start, length)
    }
  }

  private def workDone: Boolean = {
    !indexes.isEmpty && result.length == indexes.size
  }

  def getResult: ListBuffer[String] ={
    result
  }
}