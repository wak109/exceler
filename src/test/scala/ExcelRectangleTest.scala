/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.nio.file.{Paths, Files}

import ExcelLib.Converters._
import ExcelLib.Rectangle.Converters._


class ExcelRectangleSuite extends FunSuite with BeforeAndAfterEach {
  
    import ExcelRectangleSheetConversion.Helper

    val testSheet = "test"
    val testMessage = "Hello, world!!"

    override def beforeEach() {
    }
    
    override def afterEach() {
    }

    test("Cell.hasBorderBottom") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderBottom)

        cell.setBorderBottom(BorderStyle.THIN)
        assert(cell.hasBorderBottom)

        cell.setBorderBottom(BorderStyle.NONE)
        assert(! cell.hasBorderBottom)

        cell.lowerCell.setBorderTop(BorderStyle.THIN)
        assert(cell.hasBorderBottom)
    }

    test("Cell.hasBorderTop") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderTop)

        cell.setBorderTop(BorderStyle.THIN)
        assert(cell.hasBorderTop)

        cell.setBorderTop(BorderStyle.NONE)
        assert(! cell.hasBorderTop)

        cell.upperCell.setBorderBottom(BorderStyle.THIN)
        assert(cell.hasBorderTop)

        assert(sheet.cell(0,5).hasBorderTop)
    }

    test("Cell.hasBorderLeft") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderLeft)

        cell.setBorderLeft(BorderStyle.THIN)
        assert(cell.hasBorderLeft)

        cell.setBorderLeft(BorderStyle.NONE)
        assert(! cell.hasBorderLeft)

        cell.leftCell.setBorderRight(BorderStyle.THIN)
        assert(cell.hasBorderLeft)

        assert(sheet.cell(5,0).hasBorderLeft)
    }

    test("Cell.hasBorderRight") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderRight)

        cell.setBorderRight(BorderStyle.THIN)
        assert(cell.hasBorderRight)

        cell.setBorderRight(BorderStyle.NONE)
        assert(! cell.hasBorderRight)

        cell.rightCell.setBorderLeft(BorderStyle.THIN)
        assert(cell.hasBorderRight)
    }

    test("Cell.hasThickBorderRight") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderRight)

        cell.setBorderRight(BorderStyle.THIN)
        assert(!cell.hasThickBorderRight)

        cell.setBorderRight(BorderStyle.THICK)
        assert(cell.hasThickBorderRight)

        cell.setBorderRight(BorderStyle.THIN)
        assert(!cell.hasThickBorderRight)

        cell.rightCell.setBorderLeft(BorderStyle.THICK)
        assert(cell.hasThickBorderRight)
    }

    test("Cell.isOuterBorderBottom etc") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val cell = sheet.cell(5, 5)

        cell.setBorderTop(BorderStyle.NONE)
        cell.setBorderBottom(BorderStyle.THIN)
        cell.setBorderLeft(BorderStyle.NONE)
        cell.setBorderRight(BorderStyle.THIN)

        assert(cell.isOuterBorderBottom)
        assert(cell.isOuterBorderRight)

        assert(cell.getCellStyle.getBorderBottomEnum == BorderStyle.THIN)
        assert(cell.lowerCell.getCellStyle.getBorderTopEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderBottom)

        assert(cell.getCellStyle.getBorderRightEnum == BorderStyle.THIN)
        assert(cell.rightCell.getCellStyle.getBorderLeftEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderRight)
    }

    test("Cell.findTopRightFromBottomRight") {
        val sheet = (new XSSFWorkbook).sheet(testSheet)
        val topRight = sheet.cell(5, 10)
        val bottomRight = sheet.cell(10, 10)

        for {i <- (5 to 10)
                cell = sheet.cell(i, 10)
        } cell.setBorderRight(BorderStyle.THIN)

        topRight.setBorderTop(BorderStyle.THIN)
        bottomRight.setBorderBottom(BorderStyle.THIN)

        Helper.findTopRightFromBottomRight(bottomRight) match {
            case Some(c) => assert(c == topRight)
            case None => assert(false)
        }
    }

    test("Cell.findBottomLeftFromBottomRight") {
        val sheet = (new XSSFWorkbook).sheet(testSheet)
        val bottomLeft = sheet.cell(10, 5)
        val bottomRight = sheet.cell(10, 10)

        for {i <- (5 to 10)
                cell = sheet.cell(10, i)
        } cell.setBorderBottom(BorderStyle.THIN)

        bottomLeft.setBorderLeft(BorderStyle.THIN)
        bottomRight.setBorderRight(BorderStyle.THIN)

        Helper.findBottomLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(c == bottomLeft)
            case None => assert(false)
        }
    }

    test("Cell.findBottomLeftFromBottomRight: Negative Case") {
        val sheet = (new XSSFWorkbook).sheet(testSheet)
        val bottomLeft = sheet.cell(10, 5)
        val bottomRight = sheet.cell(10, 10)

        for {i <- (6 to 10)
                cell = sheet.cell(10, i)
        } cell.setBorderBottom(BorderStyle.THIN)

        bottomLeft.setBorderLeft(BorderStyle.THIN)
        bottomRight.setBorderRight(BorderStyle.THIN)

        Helper.findBottomLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }
    }

    test("Cell.findTopLeftFromBottomRight") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val topLeft = sheet.cell(5, 5)
        val topRight = sheet.cell(5, 10)
        val bottomLeft = sheet.cell(10, 5)
        val bottomRight = sheet.cell(10, 10)

        bottomLeft.setBorderLeft(BorderStyle.THIN)
        bottomLeft.setBorderBottom(BorderStyle.THIN)
        topRight.setBorderTop(BorderStyle.THIN)
        topRight.setBorderRight(BorderStyle.THIN)

        Helper.findTopLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(c.getRowIndex == 5 && c.getColumnIndex == 5)
            case None => assert(false)
        }
    }

    test("check apache poi") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val topRight = sheet.cell(5, 10)
        val bottomLeft = sheet.cell(10, 5)
        val bottomRight = sheet.cell(10, 10)

        assert(sheet.getFirstRowNum == 5)
        assert(sheet.getLastRowNum == 10)

        assert(sheet.getRow(5).getFirstCellNum == 10)
        assert(sheet.getRow(5).getLastCellNum == 11)

        assert(sheet.getRow(10).getFirstCellNum == 5)
        assert(sheet.getRow(10).getLastCellNum == 11)
    }

    test("getCellList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        val topRight = sheet.cell(5, 10)
        val bottomLeft = sheet.cell(10, 5)
        val bottomRight = sheet.cell(10, 10)

        assert(Helper.getCellList(sheet).length == 3)
    }

    test("getRectangleList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet(testSheet)
        var topLeft = sheet.cell(5, 5)
        var topRight = sheet.cell(5, 10)
        var bottomLeft = sheet.cell(10, 5)
        var bottomRight = sheet.cell(10, 10)

        bottomRight.setBorderRight(BorderStyle.THIN)
        bottomRight.setBorderBottom(BorderStyle.THIN)
        bottomLeft.setBorderLeft(BorderStyle.THIN)
        bottomLeft.setBorderBottom(BorderStyle.THIN)
        topRight.setBorderTop(BorderStyle.THIN)
        topRight.setBorderRight(BorderStyle.THIN)

        topLeft = sheet.cell(50, 50)
        topRight = sheet.cell(50, 100)
        bottomLeft = sheet.cell(100, 50)
        bottomRight = sheet.cell(100, 100)

        bottomRight.setBorderRight(BorderStyle.THIN)
        bottomRight.setBorderBottom(BorderStyle.THIN)
        bottomLeft.setBorderLeft(BorderStyle.THIN)
        bottomLeft.setBorderBottom(BorderStyle.THIN)
        topRight.setBorderTop(BorderStyle.THIN)
        topRight.setBorderRight(BorderStyle.THIN)

        val rectList = sheet.getRectangleList[ExcelTable]
        
        assert(rectList.length == 2)
    }

    test("RectDrawer mixin") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect1 = new ExcelTable(sheet, 10, 10, 20, 20)
          with RectangleBorderDraw
        val rect2 = new ExcelTable(sheet, 30, 30, 40, 40)
          with RectangleBorderDraw

        rect1.drawOuterBorder(BorderStyle.THIN)
        rect2.drawOuterBorder(BorderStyle.THIN)

        val rectList = sheet.getRectangleList[ExcelTable]
        assert(rectList.length == 2)
    }

    test("getRowList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet("test")
        val rect = new ExcelTable(sheet, 10, 10, 20, 20)
          with RectangleBorderDraw

        rect.drawOuterBorder(BorderStyle.THIN)
        rect.drawHorizontalLine(2, BorderStyle.THIN)
        rect.drawHorizontalLine(7, BorderStyle.THIN)

        rect.drawVerticalLine(2, BorderStyle.THIN)
        rect.drawVerticalLine(5, BorderStyle.THIN)
        rect.drawVerticalLine(7, BorderStyle.THIN)

        assert(rect.getRowList.length == 3)
        assert(rect.getRowList.apply(0).getColumnList.length == 4)
        assert(rect.getRowList.apply(1).getColumnList.length == 4)

        val row1 = new ExcelTable(rect.getRowList.apply(1))
          with RectangleBorderDraw
        row1.drawVerticalLine(3, BorderStyle.THIN)
        assert(rect.getRowList.apply(1).getColumnList.length == 5)
    }
}
