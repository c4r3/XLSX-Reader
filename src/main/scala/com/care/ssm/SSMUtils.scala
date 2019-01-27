package com.care.ssm

import java.io.FileInputStream
import java.util.zip.{ZipEntry, ZipInputStream}

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

    //zip file content
    val zis = new ZipInputStream(new FileInputStream(xlsxPath))

    var ze = zis.getNextEntry()
    while (ze != null) {

      val fName = ze.getName()
      if (fName.equals(required)) {
        return Some(zis)
      }
      ze = zis.getNextEntry
    }
    None : Option[ZipInputStream]
  }


}