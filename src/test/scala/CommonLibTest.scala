/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import org.scalatest.FunSuite

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._

import java.io.File
import java.nio.file.{Paths, Files}

import CommonLib.ImplicitConversions._
import CommonLib._

class CommonLibTest extends FunSuite {
  test("List splitBy") {
    assert(List(1, 2, 2, 3, 3, 3, 4).splitBy(_%2==0) ==
      List(List(1), List(2, 2), List(3, 3, 3), List(4)))
    assert(List(1, 2, 2, 3, 3, 3, 4).splitBy(_%2!=0) ==
      List(List(1), List(2, 2), List(3, 3, 3), List(4)))
  }

  test("getListOfFiles") {
    println(getListOfFiles("."))
  }
}
