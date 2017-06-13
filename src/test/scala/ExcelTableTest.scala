/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib.Converters._
import ExcelTableLib._

object ExcelTableSuite {
    def createStringEqual(s:String) = (x:String) => x == s
}

class ExcelTableSuite extends FunSuite with BeforeAndAfterEach {
  
    import ExcelTableSuite._

    val testWorkbook1 = "test1.xlsx"
    
    test("ExcelTable#getSingleValue (1x1)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")

        sheet.cell(5, 5).setCellValue("foo")
        val table = new ExcelTable(sheet, 5, 5, 5, 5)

        assert(table.getSingleValue.isDefined)
        assert(table.getSingleValue.get == "foo")

        assert(table.value.isDefined)
        assert(table.value.get == "foo")
    }

    test("ExcelTable#getSingleValue (5x5)") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")

        sheet.cell(6, 6).setCellValue("foo")
        val table = new ExcelTable(sheet, 5, 5, 10, 10)

        assert(table.getSingleValue.isDefined)
        assert(table.getSingleValue.get == "foo")

        assert(table.value.isDefined)
        assert(table.value.get == "foo")
    }

    test("ExcelTable#queryRow") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val table = new ExcelTable(sheet, 10, 10, 20, 20)
                with RectangleBorderDraw

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

        assert(table.queryRow(_ == "row1")(0).topRow == 10)
        assert(table.queryRow(
            List((x:String) => x == "row1")).apply(0).topRow == 10)
        assert(table.queryRow(_ == "row2")(0).topRow == 12)
        assert(table.queryRow(_ == "row3")(0).topRow == 17)
        assert(table.queryRow(_ == "foo").isEmpty)

        sheet.cell(10,14).setCellValue("col2")
        sheet.cell(20,16).setCellValue("bar")
        sheet.cell(10,17).setCellValue("col3")
        sheet.cell(20,17).setCellValue("bar")

        assert(table.queryColumn(_ == "col2")(0).leftCol == 12)
        assert(table.queryColumn(
            List((x:String) => x == "col2")).apply(0).leftCol == 12)
        assert(table.queryColumn(_ == "col3")(0).leftCol == 17)
        assert(table.queryColumn(_ == "bar").isEmpty)

        sheet.cell(17,17).setCellValue("hello")
    }

    test("ExcelTable#query") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.sheet("table")

        val table = new ExcelTable(sheet, 2, 1, 17, 8)
        val cell = table.query(
            List("row1", "lower").map(createStringEqual(_)),
            List("col2", "right").map(createStringEqual(_))
            )
        assert(cell(0)(0).getSingleValue.get == "lr")

        val cell2 = table.query(
            List("row1", "upper").map(createStringEqual(_)),
            List("col2", "left").map(createStringEqual(_))
            )
        assert(cell2(0)(0).getSingleValue.get == "ul")

        assert(table.getTableName.get == "test")
    }
}
