package com.caretech.ssm

import java.io.FileInputStream
import java.lang.Math.pow
import java.lang.String.valueOf
import java.util.zip.ZipInputStream


/**
  * @author Massimo Caresana
  */
object SSMUtils {

  final val content_types = "[Content_Types].xml"
  final val rels = "_rels/.rels"
  final val workbook_rels = "xl/_rels/workbook.xml.rels"
  final val workbook = "xl/workbook.xml"
  final val shared_strings = "xl/sharedStrings.xml"
  final val sheets_rels_folder = "xl/worksheets/_rels"
  final val theme_folder = "xl/theme"
  final val styles = "xl/styles.xml"
  final val sheets_folder = "xl/worksheets"
  final val core = "docProps/core.xml"
  final val printer_settings_folder = "xl/printerSettings"
  final val app = "docProps/app.xml"

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
}