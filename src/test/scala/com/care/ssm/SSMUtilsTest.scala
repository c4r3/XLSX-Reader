package com.care.ssm

import com.care.ssm.SSMUtils.SSCellType._
import com.care.ssm.SSMUtils.{SSCellType, calculateColumn, detectCellType}
import org.scalatest.flatspec.AnyFlatSpec
import org.scalatest.matchers.should.Matchers

class SSMUtilsTest extends AnyFlatSpec with Matchers {

  "column calculation" should " be correct " in {

    assert(calculateColumn("A1", 1) == 1)
    assert(calculateColumn("Z1", 1) == 26)
    assert(calculateColumn("AB1", 1) == 28)
  }

  "The zipstream " should " be correct" in {

    val result = SSMUtils.extractStream("./src/test/resources/sample_1/sample.xlsx", SSMUtils.workbook)
    result should not be null
    result.isDefined should be(true)
    result.get.available should be(1)

    val result_1 = SSMUtils.extractStream("./src/test/resources/sample_1/sample.xlsx", "not-existing-file")
    result_1 should not be null
    result_1.isDefined should be(false)
  }

  "The cell type detection " should " be correct" in {

    detectCellType("d").get should be(Date)
    detectCellType("e").get should be(Error)
    detectCellType("inlineStr").get should be(InlineString)
    detectCellType("s").get should be(SharedString)
    detectCellType("n").get should be(Double)
    detectCellType("not-existing").get should be(Long)
  }
}