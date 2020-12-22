package com.care.ssm

import java.io.FileInputStream
import java.lang.Math.pow
import java.lang.String.valueOf
import java.util.zip.ZipInputStream

/**
  * @author Massimo Caresana
  */
object SSMUtils {

  val content_types = "[Content_Types].xml"
  val rels = "_rels/.rels"
  val workbook_rels = "xl/_rels/workbook.xml.rels"
  val workbook = "xl/workbook.xml"
  val shared_strings = "xl/sharedStrings.xml"
  val sheets_rels_folder = "xl/worksheets/_rels"
  val theme_folder = "xl/theme"
  val styles = "xl/styles.xml"
  val sheets_folder = "xl/worksheets"
  val core = "docProps/core.xml"
  val printer_settings_folder = "xl/printerSettings"
  val app = "docProps/app.xml"

  def extractStream(xlsxPath: String, required: String): Option[ZipInputStream] = {

    val zis = new ZipInputStream(new FileInputStream(xlsxPath))
    var ze = zis.getNextEntry
    while (ze != null) {

      val fName = ze.getName
      if (fName.equals(required)) {
        return Some(zis)
      }
      ze = zis.getNextEntry
    }
    None: Option[ZipInputStream]
  }

  val lettersNum: Int = 'Z'.toInt - 'A'.toInt + 1

  def calculateColumn(str: String, rowNum: Int): Int = {

    //Remove rowNum "AB1" -> "AB"
    val chars = str.replace(valueOf(rowNum), "").toCharArray

    chars.reverse.zipWithIndex.map(
      pair => (pair._1.toInt - 'A'.toInt + 1) * pow(lettersNum, pair._2).toInt
    ).sum
  }

  def toInt(s: String): Option[Int] = {
    try {
      Some(s.toInt)
    } catch {
      case ex : Exception =>
        ex.printStackTrace()
        println(s"Error toInt parsing $s")
        None
    }
  }

  //FIXME portalo a Either Try(s.toDouble).toEither(...)
  def toDouble(s: String): Option[Double] = {
    try {
      Some(s.toDouble)
    } catch {
      case ex : Exception =>
        ex.printStackTrace()
        println(s"Error toDouble parsing $s")
        None
    }
  }

  def toLong(s: String): Option[Long] = {
    try {
      Some(s.toLong)
    } catch {
      case ex : Exception =>
        ex.printStackTrace()
        println(s"Error toLong parsing $s")
        None
    }
  }
}