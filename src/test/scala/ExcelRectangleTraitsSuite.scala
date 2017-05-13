/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.nio.file.{Paths, Files}

import ExcelLib._
import ExcelRectangle._

class ExcelRectangleTraitsSuite extends FunSuite with BeforeAndAfterEach {
  
    override def beforeEach() {
    }
    
    override def afterEach() {
    }

    test("ExcelRectangleDraw mixin") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect1 = new ExcelRectangle(sheet, 10, 10, 20, 20) with ExcelRectangleDraw
        val rect2 = new ExcelRectangle(sheet, 30, 30, 40, 40) with ExcelRectangleDraw

        rect1.drawOuterBorder(BorderStyle.THIN)
        rect2.drawOuterBorder(BorderStyle.THIN)

        val rectList = sheet.getRectangleList
        assert(rectList.length == 2)
    }

    test("getInnerRectangleList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect = new ExcelRectangle(sheet, 10, 10, 20, 20) with ExcelRectangleDraw

        rect.drawOuterBorder(BorderStyle.THIN)
        rect.drawHorizontalLine(2, BorderStyle.THIN)
        rect.drawHorizontalLine(7, BorderStyle.THIN)

        rect.drawVerticalLine(2, BorderStyle.THIN)
        rect.drawVerticalLine(5, BorderStyle.THIN)
        rect.drawVerticalLine(7, BorderStyle.THIN)

        assert(rect.getInnerRectangleList.length == 3)
        assert(rect.getInnerRectangleList()(0).length == 4)
        assert(rect.getInnerRectangleList()(1).length == 4)
        assert(rect.getInnerRectangleList()(2).length == 4)
    }
}
