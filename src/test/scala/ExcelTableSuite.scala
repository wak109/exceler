/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


import java.nio.file.{Paths, Files}

import ExcelLib._
import ExcelTable._

class ExcelTableSuite extends FunSuite with BeforeAndAfterEach {
  
    val testSheet = "test"
    val testMessage = "Hello, world!!"

    override def beforeEach() {
    }
    
    override def afterEach() {
    }

    test("Cell.hasBorderBottom_") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderBottom_)

        cell.setBorderBottom_(BorderStyle.THIN)
        assert(cell.hasBorderBottom_)

        cell.setBorderBottom_(BorderStyle.NONE)
        assert(! cell.hasBorderBottom_)

        cell.lowerCell_.setBorderTop_(BorderStyle.THIN)
        assert(cell.hasBorderBottom_)
    }

    test("Cell.hasBorderTop_") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderTop_)

        cell.setBorderTop_(BorderStyle.THIN)
        assert(cell.hasBorderTop_)

        cell.setBorderTop_(BorderStyle.NONE)
        assert(! cell.hasBorderTop_)

        cell.upperCell_.setBorderBottom_(BorderStyle.THIN)
        assert(cell.hasBorderTop_)

        assert(sheet.cell_(0,5).hasBorderTop_)
    }

    test("Cell.hasBorderLeft_") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderLeft_)

        cell.setBorderLeft_(BorderStyle.THIN)
        assert(cell.hasBorderLeft_)

        cell.setBorderLeft_(BorderStyle.NONE)
        assert(! cell.hasBorderLeft_)

        cell.leftCell_.setBorderRight_(BorderStyle.THIN)
        assert(cell.hasBorderLeft_)

        assert(sheet.cell_(5,0).hasBorderLeft_)
    }

    test("Cell.hasBorderRight_") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)
        var style = workbook.createCellStyle

        assert(! cell.hasBorderRight_)

        cell.setBorderRight_(BorderStyle.THIN)
        assert(cell.hasBorderRight_)

        cell.setBorderRight_(BorderStyle.NONE)
        assert(! cell.hasBorderRight_)

        cell.rightCell_.setBorderLeft_(BorderStyle.THIN)
        assert(cell.hasBorderRight_)
    }

    test("Cell.isOuterBorderBottom_ etc") {
        val workbook = new XSSFWorkbook
        val sheet = workbook.sheet_(testSheet)
        val cell = sheet.cell_(5, 5)

        cell.setBorderTop_(BorderStyle.NONE)
        cell.setBorderBottom_(BorderStyle.THIN)
        cell.setBorderLeft_(BorderStyle.NONE)
        cell.setBorderRight_(BorderStyle.THIN)

        assert(cell.isOuterBorderBottom_)
        assert(cell.isOuterBorderRight_)

        assert(cell.getCellStyle.getBorderBottomEnum == BorderStyle.THIN)
        assert(cell.lowerCell_.getCellStyle.getBorderTopEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderBottom_)

        assert(cell.getCellStyle.getBorderRightEnum == BorderStyle.THIN)
        assert(cell.rightCell_.getCellStyle.getBorderLeftEnum == BorderStyle.NONE)
        assert(cell.isOuterBorderRight_)
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

    test("getExcelTableList") {
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

        val tableList = getExcelTableList(sheet)
        
        assert(tableList.length == 2)
    }
}
