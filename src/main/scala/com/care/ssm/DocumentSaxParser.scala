package com.care.ssm

import java.util.zip.ZipInputStream

import javax.xml.parsers.SAXParserFactory


class DocumentSaxParser {


  val factory =  SAXParserFactory.newInstance
  val parser = factory.newSAXParser
  val handler = new SaxHandler

  def readSheet(sheetName: String): Unit ={



  }

  def lookupSheetIdByName(xlsxPath: String, sheetName: String): Unit ={

    val zis = SSMUtils.extractStream(xlsxPath, SSMUtils.workbook)

    if (zis.isDefined){
      println("check")
      parser.parse(zis.get, handler)

    } else {
      println("Something wrong during zip stream extraction")
    }
  }

  def read(zis : ZipInputStream): Unit ={
    parser.parse(zis,handler)
  }
}