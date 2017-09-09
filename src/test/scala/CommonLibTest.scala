/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import org.scalatest.FunSuite

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._

import java.io.File
import java.nio.file.{Paths, Files}

import exceler.common._
import CommonLib.ImplicitConversions._
import CommonLib._

class CommonLibTest extends FunSuite {
  test("List splitBy") {
    assert(List(1, 2, 2, 3, 3, 3, 4).splitBy(_%2==0) ==
      List(List(1), List(2, 2), List(3, 3, 3), List(4)))
    assert(List(1, 2, 2, 3, 3, 3, 4).splitBy(_%2!=0) ==
      List(List(1), List(2, 2), List(3, 3, 3), List(4)))
  }

  test("List blockingBy") {
    assert(List(1,2,2,4,3,1,2,4,2,1,2,4,3,4,4,3).blockingBy(_%2==0) ==
      List(List(2,2,4), List(2,4,2), List(2,4), List(4,4)))
    assert(List(2,2,4,3,1,2,4,2,1,2,4,3,4,4,3).blockingBy(_%2==0) ==
      List(List(2,2,4), List(2,4,2), List(2,4), List(4,4)))
    assert(List(1,2,2,4,3,1,2,4,2,1,2,4,3,4,4).blockingBy(_%2==0) ==
      List(List(2,2,4), List(2,4,2), List(2,4), List(4,4)))
    assert(List(2,2,4,3,1,2,4,2,1,2,4,3,4,4).blockingBy(_%2==0) ==
      List(List(2,2,4), List(2,4,2), List(2,4), List(4,4)))
  }

  test("getListOfFiles") {
    println(getListOfFiles("."))
  }
}
