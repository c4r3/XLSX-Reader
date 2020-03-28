package com.care.ssm

import com.care.ssm.SSMUtils.calculateColumn
import org.scalatest.flatspec.AnyFlatSpec

object SSMUtilsTest extends AnyFlatSpec {

  "column calculation" should " be correct " in {

    assert(calculateColumn("A1", 1) == 1)
    assert(calculateColumn("Z1", 1) == 26)
    assert(calculateColumn("AB1", 1) == 28)
  }
}