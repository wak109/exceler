/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib.ImplicitConversions._


class ExcelTableLibTest
  extends FunSuite with ExcelLibResource
  with ExcelTableSheetFunction {

  test("TableQueryTraitImpl.getValue") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    val func = new TableQueryTraitImpl {}
    assert(func.getValue(sheet, 2, 1) == Some("test3"))
    assert(func.getValue(sheet, 2, 6) == Some("test3"))
    assert(func.getValue(sheet, 11, 4, 12, 5) == Some("what is it?"))

  }

  test("TableQuery2") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    val tableQuery = new TableQuery2[Rect](new Rect(sheet,2,1,18,8))
  }

  test("collectTopLeft") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    assert(ExcelTableLib.collectTopLeft(sheet,2,1,18,8) == List((2,1), (3,1), (3,3), (3,5), (3,7), (5,3), (5,4), (5,5), (5,6), (7,1), (7,2), (7,3), (7,4), (7,5), (7,6), (7,7), (7,8), (9,2), (9,3), (9,4), (9,6), (9,7), (9,8), (11,2), (11,3), (11,6), (11,7), (11,8), (13,2), (13,3), (13,4), (13,5), (13,6), (13,7), (13,8), (15,1), (15,2), (15,3), (15,4), (15,5), (15,6), (15,7), (15,8), (17,2), (17,3), (17,4), (17,5), (17,6), (17,7), (17,8))) 
  }

  test("convertTopLeftToRect") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    assert(ExcelTableLib.getBottomRight(sheet, 2, 1, 18, 8) == (2,8))
    assert(ExcelTableLib.getBottomRight(sheet, 9, 4, 18, 8) == (12,5))
    assert(ExcelTableLib.getBottomRight(sheet, 17, 8, 18, 8) == (18,8)) 
  }

  test("getRowList") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    assert(ExcelTableLib.getRowList(
      ExcelTableLib.collectTopLeft(sheet,2,1,18,8)) == 
        List(2, 3, 5, 7, 9, 11, 13, 15, 17))
  }

  test("getColumnList") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")

    assert(ExcelTableLib.getColumnList(
      ExcelTableLib.collectTopLeft(sheet,2,1,18,8)) == 
        List(1,2,3,4,5,6,7,8))
  }

  test("getRowSpan,getColumnSpan") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("table2")
    val rowList = ExcelTableLib.getRowList(
      ExcelTableLib.collectTopLeft(sheet,2,1,18,8))
    val columnList = ExcelTableLib.getColumnList(
      ExcelTableLib.collectTopLeft(sheet,2,1,18,8))

    assert(ExcelTableLib.getRowSpan(2,2,rowList) == List(0))
    assert(ExcelTableLib.getRowSpan(3,6,rowList) == List(1,2))

    assert(ExcelTableLib.getRowSpan(1,4,columnList) == List(0,1,2,3))
  }
}
