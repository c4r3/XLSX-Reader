package com.care.ssm

import java.io.FileInputStream
import java.util.zip.ZipInputStream

import com.care.ssm.SSMUtils.SSCellType.SSCellType

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
    None: Option[ZipInputStream]
  }

  val lettersNum = 'Z'.toInt - 'A'.toInt + 1

  def calculateColumn(str: String, rowNum: Int): Int = {

    //Remove rowNum "AB1" -> "AB"
    val chars = str.replace(String.valueOf(rowNum), "").toCharArray

    val result: Int = chars.reverse.zipWithIndex.map(
      pair => (pair._1.toInt - 'A'.toInt + 1) * Math.pow(lettersNum, pair._2).toInt
    ).sum
    return result
  }

  def toInt(s: String): Option[Int] = {
    try {
      Some(s.toInt)
    } catch {
      case e: Exception => None
    }
  }

  def toDouble(s: String): Option[Double] = {
    try {
      Some(s.toDouble)
    } catch {
      case e: Exception => None
    }
  }

  //TODO metodo per effettuare la detection del tipo della cella: sarebbe un valore di un enum.
  //TODO In tal modo ogni cella avrà solo un int per identificare il tipo e non un int e una stringa
  //TODO si risparmia memoria
  /**
    *
    * https://msdn.microsoft.com/library/office/documentformat.openxml.spreadsheet.cell.aspx
    * <pre>
    * -------------------------------------------------------------------------
    * Enumeration Value          Description
    * -------------------------------------------------------------------------
    * b (Boolean)                Cell containing a boolean.
    * d (Date)                   Cell contains a date in the ISO 8601 format.
    * e (Error)                  Cell containing an error.
    * inlineStr (Inline String)  Cell containing an (inline) rich string, i.e.,
    * one not in the shared string table. If this
    * cell type is used, then the cell value is in
    * the is element rather than the v element in
    * the cell (c element).
    * n (Number)                 Cell containing a number.
    * s (Shared String)          Cell containing a shared string.
    * str (String)               Cell containing a formula string
    * </pre>
    * @param style
    * @param rawType
    * @return
    */
  def detectCellType(style: String, rawType: String): SSCellType = {

    rawType match {

      case "d" => SSCellType.Date
      case "e" => SSCellType.Error
      case "inlineStr" => SSCellType.InlineString
      case "s" => SSCellType.SharedString
      case "n" => SSCellType.Double
      case _ => {
        style match {
          case some =>
          case _ => SSCellType.Long
        }
      }
    }
    SSCellType.Unknown
  }

  object SSCellType extends Enumeration {
    type SSCellType = Value
    val String, SharedString, InlineString, Long, Date, Double, Error, Unknown = Value
  }
}