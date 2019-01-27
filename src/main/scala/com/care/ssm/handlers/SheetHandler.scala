package com.care.ssm.handlers

import java.util

import com.care.ssm.handlers.SheetHandler.SSCell
import org.xml.sax.Attributes
import org.xml.sax.helpers.DefaultHandler

/**
  * Handler for Shared Strings File
  * @param indexes
  */
class SheetHandler(indexes: Set[Int] = Set[Int]()) extends DefaultHandler{

  var result = new util.ArrayList[SSCell]

  //Target and value tags
  val targetTag = "c"
  val valueTag = "v"
  //Attributes tags
  val attr1 = "r"
  val attr2 = "s"
  val attr3 = "t"

  var cellEnded = false

  var valueTagEnded = false

  //Temporary variables
  var cellRow = ""
  var cellStyle = ""
  var cellType = ""
  var cellValue = ""

  override def startElement(uri: String, localName: String, qName: String, attributes: Attributes): Unit = {

    if(workDone) return

    if(targetTag.equals(qName)) {
      //Starting "c" tag, extraction of the attributes
      cellRow = attributes.getValue(attr1)
      cellStyle = attributes.getValue(attr2)
      cellType = attributes.getValue(attr3)
    }
  }

  override def endElement(uri: String, localName: String, qName: String): Unit = {

    if (workDone) return

    if(valueTag.equals(qName)) valueTagEnded = true

    if(targetTag.equals(qName)){
      //Closing event of target tag
      val currentCell = new SSCell(cellRow, cellStyle, cellType, cellValue)

    }

  }

  override def characters(ch: Array[Char], start: Int, length: Int): Unit = {

    if (workDone) return

    if(valueTagEnded) {
      //Flushing buffer
      cellValue = new String(ch, start, length)
      val currentCell = new SSCell(cellRow, cellStyle, cellType, cellValue)

      result.add(currentCell)

      valueTagEnded = false
      cellRow = ""
      cellStyle = ""
      cellType = ""
      cellValue = ""
    }
  }

  private def workDone: Boolean = {
    false //TODO da completare
  }

  def getResult: util.ArrayList[SSCell] ={
    result
  }
}

object SheetHandler {

  //TODO devono essere specializzate le classi con i vari tipi in overloading,verrà fatto dopo che si capisce
  //TODO la logica per determinare l'entità dagli attributi
  case class SSCell(row: String, style: String, ctype: String, value: String)
}