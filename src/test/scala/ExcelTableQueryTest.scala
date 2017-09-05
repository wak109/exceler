/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib.ImplicitConversions._


class ExcelTableQueryTest
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

}
