/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.nio.file.{Paths, Files}

import ExcelLib._
import ExcelRectangle._
import ExcelRectangle.Helper._

class ExcelRectangleSuite extends FunSuite with BeforeAndAfterEach {
  
    val testSheet = "test"
    val testMessage = "Hello, world!!"

    override def beforeEach() {
    }
    
    override def afterEach() {
    }

    test("Cell.hasBorderBottom") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderBottom)

        cell.setBorderBottom_(BorderStyle.THIN)
        assert(cell.hasBorderBottom)

        cell.setBorderBottom_(BorderStyle.NONE)
        assert(! cell.hasBorderBottom)

        cell.lowerCell_.setBorderTop_(BorderStyle.THIN)
        assert(cell.hasBorderBottom)
    }

    test("Cell.hasBorderTop") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderTop)

        cell.setBorderTop_(BorderStyle.THIN)
        assert(cell.hasBorderTop)

        cell.setBorderTop_(BorderStyle.NONE)
        assert(! cell.hasBorderTop)

        cell.upperCell_.setBorderBottom_(BorderStyle.THIN)
        assert(cell.hasBorderTop)

        assert(sheet.cell_(0,5).hasBorderTop)
    }

    test("Cell.hasBorderLeft") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderLeft)

        cell.setBorderLeft_(BorderStyle.THIN)
        assert(cell.hasBorderLeft)

        cell.setBorderLeft_(BorderStyle.NONE)
        assert(! cell.hasBorderLeft)

        cell.leftCell_.setBorderRight_(BorderStyle.THIN)
        assert(cell.hasBorderLeft)

        assert(sheet.cell_(5,0).hasBorderLeft)
    }

    test("Cell.hasBorderRight") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderRight)

        cell.setBorderRight_(BorderStyle.THIN)
        assert(cell.hasBorderRight)

        cell.setBorderRight_(BorderStyle.NONE)
        assert(! cell.hasBorderRight)

        cell.rightCell_.setBorderLeft_(BorderStyle.THIN)
        assert(cell.hasBorderRight)
    }

    test("Cell.isOuterBorderBottom etc") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)

        cell.setBorderTop_(BorderStyle.NONE)
        cell.setBorderBottom_(BorderStyle.THIN)
        cell.setBorderLeft_(BorderStyle.NONE)
        cell.setBorderRight_(BorderStyle.THIN)

        assert(cell.isOuterBorderBottom)
        assert(cell.isOuterBorderRight)

        assert(cell.getCellStyle.getBorderBottomEnum == BorderStyle.THIN)
        assert(cell.lowerCell_.getCellStyle.getBorderTopEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderBottom)

        assert(cell.getCellStyle.getBorderRightEnum == BorderStyle.THIN)
        assert(cell.rightCell_.getCellStyle.getBorderLeftEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderRight)
    }

    test("Cell.findTopRightFromBottomRight") {
        val sheet = (new XSSFWorkbook).sheet_(testSheet)
        val topRight = sheet.cell_(5, 10)
        val bottomRight = sheet.cell_(10, 10)

        for {i <- (5 to 10)
                cell = sheet.cell_(i, 10)
        } cell.setBorderRight_(BorderStyle.THIN)

        topRight.setBorderTop_(BorderStyle.THIN)
        bottomRight.setBorderBottom_(BorderStyle.THIN)

        findTopRightFromBottomRight(bottomRight) match {
            case Some(c) => assert(c == topRight)
            case None => assert(false)
        }
    }

    test("Cell.findBottomLeftFromBottomRight") {
        val sheet = (new XSSFWorkbook).sheet_(testSheet)
        val bottomLeft = sheet.cell_(10, 5)
        val bottomRight = sheet.cell_(10, 10)

        for {i <- (5 to 10)
                cell = sheet.cell_(10, i)
        } cell.setBorderBottom_(BorderStyle.THIN)

        bottomLeft.setBorderLeft_(BorderStyle.THIN)
        bottomRight.setBorderRight_(BorderStyle.THIN)

        findBottomLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(c == bottomLeft)
            case None => assert(false)
        }
    }

    test("Cell.findBottomLeftFromBottomRight: Negative Case") {
        val sheet = (new XSSFWorkbook).sheet_(testSheet)
        val bottomLeft = sheet.cell_(10, 5)
        val bottomRight = sheet.cell_(10, 10)

        for {i <- (6 to 10)
                cell = sheet.cell_(10, i)
        } cell.setBorderBottom_(BorderStyle.THIN)

        bottomLeft.setBorderLeft_(BorderStyle.THIN)
        bottomRight.setBorderRight_(BorderStyle.THIN)

        findBottomLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }
    }

    test("Cell.findTopLeftFromBottomRight") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val topLeft = sheet.cell_(5, 5)
        val topRight = sheet.cell_(5, 10)
        val bottomLeft = sheet.cell_(10, 5)
        val bottomRight = sheet.cell_(10, 10)

        bottomLeft.setBorderLeft_(BorderStyle.THIN)
        bottomLeft.setBorderBottom_(BorderStyle.THIN)
        topRight.setBorderTop_(BorderStyle.THIN)
        topRight.setBorderRight_(BorderStyle.THIN)

        findTopLeftFromBottomRight(bottomRight) match {
            case Some(c) => assert(c.getRowIndex == 5 && c.getColumnIndex == 5)
            case None => assert(false)
        }
    }

    test("check apache poi") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val topRight = sheet.cell_(5, 10)
        val bottomLeft = sheet.cell_(10, 5)
        val bottomRight = sheet.cell_(10, 10)

        assert(sheet.getFirstRowNum == 5)
        assert(sheet.getLastRowNum == 10)

        assert(sheet.getRow(5).getFirstCellNum == 10)
        assert(sheet.getRow(5).getLastCellNum == 11)

        assert(sheet.getRow(10).getFirstCellNum == 5)
        assert(sheet.getRow(10).getLastCellNum == 11)
    }

    test("getCellList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val topRight = sheet.cell_(5, 10)
        val bottomLeft = sheet.cell_(10, 5)
        val bottomRight = sheet.cell_(10, 10)

        assert(getCellList(sheet).length == 3)
    }

    test("getRectangleList") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        var topLeft = sheet.cell_(5, 5)
        var topRight = sheet.cell_(5, 10)
        var bottomLeft = sheet.cell_(10, 5)
        var bottomRight = sheet.cell_(10, 10)

        bottomRight.setBorderRight_(BorderStyle.THIN)
        bottomRight.setBorderBottom_(BorderStyle.THIN)
        bottomLeft.setBorderLeft_(BorderStyle.THIN)
        bottomLeft.setBorderBottom_(BorderStyle.THIN)
        topRight.setBorderTop_(BorderStyle.THIN)
        topRight.setBorderRight_(BorderStyle.THIN)

        topLeft = sheet.cell_(50, 50)
        topRight = sheet.cell_(50, 100)
        bottomLeft = sheet.cell_(100, 50)
        bottomRight = sheet.cell_(100, 100)

        bottomRight.setBorderRight_(BorderStyle.THIN)
        bottomRight.setBorderBottom_(BorderStyle.THIN)
        bottomLeft.setBorderLeft_(BorderStyle.THIN)
        bottomLeft.setBorderBottom_(BorderStyle.THIN)
        topRight.setBorderTop_(BorderStyle.THIN)
        topRight.setBorderRight_(BorderStyle.THIN)

        val rectList = sheet.getRectangleList
        
        assert(rectList.length == 2)
    }
}
