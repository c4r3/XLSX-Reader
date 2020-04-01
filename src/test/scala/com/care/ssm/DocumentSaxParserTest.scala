package com.care.ssm

import com.care.ssm.handlers.StyleHandler.SSCellStyle
import org.scalatest.PrivateMethodTester
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

import scala.collection.mutable.ListBuffer

class DocumentSaxParserTest extends AnyFlatSpec with PrivateMethodTester with Matchers {

  "inner zip sheet file path " should " be correct " in {

    val parser = new DocumentSaxParser
    val method: PrivateMethod[String] = PrivateMethod[String](Symbol("buildInnerZipSheetFilePath"))

    "xl/worksheets/sheet1.xml" should equal (parser invokePrivate method("xl/worksheets", "1"))
    "xl/worksheets/sheet123.xml" should equal (parser invokePrivate method("xl/worksheets", "123"))
  }

  "lookup cells style " should " be correct " in {

    val parser = new DocumentSaxParser
    val method = PrivateMethod[ListBuffer[SSCellStyle]](Symbol("lookupCellsStyles"))
    val result: ListBuffer[SSCellStyle] = parser invokePrivate method("./src/test/resources/sample_1/sample.xlsx")

    result should have size 5
    result.head should be (null)
    result(1) should not be null
    result(1) should be (SSCellStyle(164, "\"$\"#,##0"))
    result(2) should be (null)
    result(3) should be (null)
    result(4) should not be null
    result(4) should be (SSCellStyle(1, "\"$\"#,##0"))
  }

  "lookup shared string " should " be correct " in {

    val parser = new DocumentSaxParser
    val method = PrivateMethod[ListBuffer[String]](Symbol("lookupSharedString"))
    val result: ListBuffer[String] = parser invokePrivate method("./src/test/resources/sample_1/sample.xlsx", Set[Int]())

    result should have size 138

    val expected = ListBuffer[String]("2000","2007","2008","2009","2010","2011","2015","2016","2017","2018",
      "2021","2022","2023","2024","2025","2026","2027","2028","2029","2030","2031","2032","2034","2037","2038","2039",
      "2040","2041","2042","2043","2044","2045","2046","2047","2048","2049","2060","2061","2062","2063","2064","2065",
      "2066","2067","2068","2069","2070","2071","2072","2073","2086","2087","2089","2090","2092","2093","2094","2095",
      "2096","2099","2100","2110","2111","2112","2114","2115","2116","2117","2118","2119","2121","2122","2125","2127",
      "2128","2130","2131","2132","2133","2134","2135","2136","2137","2141","2142","2143","2144","2145","2146","2147",
      "2151","2152","2153","2160","2161","2162","2163","2164","2165","2166","2170","2190","2191","2192","2193","2194",
      "2195","2196","2197","2198","2199","2200","2203","2204","2205","2206","2050","2006","2033","2052","2088","2091",
      "2113","2109","2138","2139","2150","2123","2140","2129","Postcode","Sales_Rep_ID","Sales_Rep_Name","John",
      "Ashish","Jane","Year","Value")

    result should be (expected)

    val result_1: ListBuffer[String] = parser invokePrivate method("./src/test/resources/sample_1/sample.xlsx",
      Set[Int](0,1,2))

    result_1 should have size 3
    result_1 should be (List[String]("2000","2007","2008"))

    val result_2: ListBuffer[String] = parser invokePrivate method("./src/test/resources/sample_1/sample.xlsx",
      Set[Int](129,130,131,132,133,134,135,136,137))

    result_2 should have size 9
    result_2 should be (List[String]("2129","Postcode","Sales_Rep_ID","Sales_Rep_Name","John","Ashish","Jane",
      "Year","Value"))
  }
}