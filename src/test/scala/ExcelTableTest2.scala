/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib.ImplicitConversions._


class ExcelTableSuite2 extends FunSuite with ExcelLibResource {
  
    test("new TableQueryImpl") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")

        val table:TableQueryImpl = new RectangleImpl(sheet, 10, 10, 20, 20)

        assert(table.rowList.length == 1)
        assert(table.colList.length == 1)
    }

    test("TableCellImpl (1x1)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")

        val cell = new TableCellImpl(new RectangleImpl(sheet, 5, 5, 5, 5))

        assert(cell.getSingleValue.isEmpty)
        sheet.cell(5, 5).setCellValue("foo")
        assert(cell.getSingleValue.isDefined)

        assert(cell.value == "foo")
    }

    test("TableCellImpl (5x5)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")

        val cell = new TableCellImpl(new RectangleImpl(sheet, 5, 5, 9, 9))

        assert(cell.getSingleValue.isEmpty)
        sheet.cell(7, 7).setCellValue("foo")
        assert(cell.getSingleValue.isDefined)

        assert(cell.value == "foo")
    }
}
