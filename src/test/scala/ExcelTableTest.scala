/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib.ImplicitConversions._


class ExcelTableSuite2 extends FunSuite with ExcelLibResource {
  
    def createStringEqual(s:String) = (x:String) => x == s

    test("new TableQueryImpl") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val table = new TableQueryImpl(sheet, 10, 10, 20, 20)

        assert(table.rowList.length == 1)
        assert(table.colList.length == 1)
    }

    test("TableCellImpl (1x1)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect = ExcelRectangle(sheet, 5, 5, 5, 5)

        assert(ExcelRectangle.function.getValue(rect).isEmpty)
        sheet.cell(5, 5).setCellValue("foo")
        assert(ExcelRectangle.function.getValue(rect).isDefined)

        // assert(cell.value == "foo")
    }

    test("TableCellImpl (5x5)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect = ExcelRectangle(sheet, 5, 5, 9, 9)

        assert(ExcelRectangle.function.getValue(rect).isEmpty)
        sheet.cell(7, 7).setCellValue("foo")
        assert(ExcelRectangle.function.getValue(rect).isDefined)

        //assert(cell.value == "foo")
    }

    test("TableQueryImpl.queryRow") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val table = new TableQueryImpl(sheet, 10, 10, 20, 20)

        table.drawOuterBorder(BorderStyle.THIN)
        table.drawHorizontalLine(2, BorderStyle.THIN)
        table.drawHorizontalLine(5, BorderStyle.THIN)
        table.drawHorizontalLine(7, BorderStyle.THIN)

        table.drawVerticalLine(2, BorderStyle.THIN)
        table.drawVerticalLine(5, BorderStyle.THIN)
        table.drawVerticalLine(7, BorderStyle.THIN)

        sheet.cell(10,10).setCellValue("row1")
        sheet.cell(14,10).setCellValue("row2")
        sheet.cell(16,20).setCellValue("foo")
        sheet.cell(17,10).setCellValue("row3")
        sheet.cell(17,20).setCellValue("foo")

        assert(table.queryRow(_ == "row1").apply(0).top == 10)
        assert(table.queryRow(
            List((x:String) => x == "row1")).apply(0).top == 10)
        assert(table.queryRow(_ == "row2").apply(0).top == 12)
        assert(table.queryRow(_ == "row3").apply(0).top == 17)
        assert(table.queryRow(_ == "foo").isEmpty)

        sheet.cell(10,14).setCellValue("col2")
        sheet.cell(20,16).setCellValue("bar")
        sheet.cell(10,17).setCellValue("col3")
        sheet.cell(20,17).setCellValue("bar")

        assert(table.queryColumn(_ == "col2")(0).left == 12)
        assert(table.queryColumn(
            List((x:String) => x == "col2")).apply(0).left == 12)
        assert(table.queryColumn(_ == "col3")(0).left == 17)
        assert(table.queryColumn(_ == "bar").isEmpty)

        sheet.cell(17,17).setCellValue("hello")
    }

    test("getTableName") {

        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.sheet("table")
        val table = new TableQueryImpl(sheet, 2, 1, 17, 8)
        val cell = table.query(
            List("row1", "lower").map(createStringEqual(_)),
            List("col2", "right").map(createStringEqual(_))
            )

        assert(ExcelRectangle.function.getValue(cell(0)(0)).get == "lr")

        val cell2 = table.query(
            List("row1", "upper").map(createStringEqual(_)),
            List("col2", "left").map(createStringEqual(_))
            )

        assert(ExcelRectangle.function.getValue(cell2(0)(0)).get == "ul")

        assert(ExcelRectangle.function.getTableName(table)._1.get == "test")
    }

    test("StackedTable") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.sheet("stack")
        val table = new TableQueryImpl(sheet, 1, 1, 18, 12)

        assert(table.rowList.length == 17)
        assert(table.colList.length == 3)

        assert(ExcelRectangle.function.getHeadCol(table.rowList(0))._2 != None)
        assert(ExcelRectangle.function.getHeadCol(table.rowList(5))._2 == None)
        assert(ExcelRectangle.function.getHeadCol(table.rowList(11))._2 == None)
        assert(ExcelRectangle.function.getHeadCol(table.rowList(12))._2 != None)

        val cell1 = table.query(
            List("separator1", "row1").map(createStringEqual(_)),
            List("col2").map(createStringEqual(_))
           )
        assert(ExcelRectangle.function.getValue(cell1(0)(0)).get == "val12-1")

        val cell2 = table.query(
            List("separator2", "row1").map(createStringEqual(_)),
            List("col3").map(createStringEqual(_))
           )
        assert(ExcelRectangle.function.getValue(cell2(0)(0)).get == "val13-2")

        val cell3 = table.query(
            List("row1").map(createStringEqual(_)),
            List("col3").map(createStringEqual(_))
           )
        assert(ExcelRectangle.function.getValue(cell3(0)(0)).get == "val13")
    }
}
