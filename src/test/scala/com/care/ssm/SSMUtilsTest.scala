package com.care.ssm

import org.scalatest.FlatSpec

object SSMUtilsTest extends FlatSpec {

  "column calculation" should " be correct " in {

    assert(SSMUtils.calculateColumn("A1", 1) == 1)
    assert(SSMUtils.calculateColumn("Z1", 1) == 26)
    assert(SSMUtils.calculateColumn("AB1", 1) == 28)

  }
}