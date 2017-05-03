/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import org.apache.poi.ss.usermodel._

import java.nio.file.{Paths, Files}

import ExcelImplicits._

class ExcelerSuite extends FunSuite with BeforeAndAfterEach {
  
    val testSheet = "test"
    val testMessage = "Hello, world!!"

    override def beforeEach() {
    }
    
    override def afterEach() {
    }
    
    //////////////////////////////////////////////////////////////// 
    // Workbook Tests
    //
    test("Workbook.saveAs_") {
        val testFile = "Test_Workbook_saveAs.xlsx"

        val workbook = ExcelerWorkbook.create()
        workbook.saveAs_(testFile)
        assert(Files.exists(Paths.get(testFile)))

        Files.deleteIfExists(Paths.get(testFile))
    }

    test("Workbook.getSheet_") {
        val workbook = ExcelerWorkbook.create()

        workbook.getSheet_(testSheet) match {
            case Some(_) => assert(false)
            case None => assert(true)
        }

        workbook.createSheet_(testSheet) match {
            case Success(_) => assert(true)
            case Failure(e) => assert(false)
        }

        workbook.getSheet_(testSheet) match {
            case Some(_) => assert(true)
            case None => assert(false)
        }

        /* Can't create sheets of same name */
        workbook.createSheet_(testSheet) match {
            case Success(_) => assert(false)
            case Failure(e) => assert(true)
        }
    }

    test("Workbook.sheet_") {
        val workbook = ExcelerWorkbook.create()

        Try(workbook.sheet_(testSheet)) match {
            case Success(s) => assert(s.getSheetName() == testSheet)
            case Failure(e) => assert(false)
        }

        Try(workbook.sheet_(testSheet)) match {
            case Success(s) => assert(s.getSheetName() == testSheet)
            case Failure(e) => assert(false)
        }
        assert(workbook.sheet_(testSheet) == workbook.getSheet(testSheet))
    }

    test("Workbook.findCellStyle_") {
        val workbook = ExcelerWorkbook.create()
        val cellStyle = workbook.sheet_(testSheet).cell_(5, 5).getCellStyle
        val tuple = cellStyle.toTuple_

        assert(workbook.findCellStyle_(tuple).nonEmpty)

        assert(workbook.findCellStyle_(
            (tuple._1, tuple._2, tuple._3, BorderStyle.THIN, tuple._5,
            tuple._6, tuple._7, tuple._8, tuple._9, tuple._10)
            ).isEmpty)

        var newStyle = workbook.createCellStyle
        newStyle.cloneStyleFrom(cellStyle)
        newStyle.setBorderRight(BorderStyle.THIN)

        assert(workbook.findCellStyle_(
            (tuple._1, tuple._2, tuple._3, BorderStyle.THIN, tuple._5,
            tuple._6, tuple._7, tuple._8, tuple._9, tuple._10)
            ).nonEmpty)
    }


    //////////////////////////////////////////////////////////////// 
    // Sheet Tests
    //
    test("Sheet.getRow_,Sheet.row_") {
        val workbook = ExcelerWorkbook.create()
        val sheet = workbook.sheet_(testSheet)

        sheet.getRow_(0) match {
            case Some(r) => assert(false)
            case None => assert(true)
        }

        val row = sheet.row_(0)

        sheet.getRow_(0) match {
            case Some(r) => assert(true)
            case None => assert(false)
        }
    }

    test("Sheet.getCell_") {
        val sheet = ExcelerWorkbook.create().sheet_(testSheet)

        sheet.getCell_(0, 0) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }

        val cell = sheet.cell_(0, 0)

        sheet.getCell_(0, 0) match {
            case Some(c) => assert(true)
            case None => assert(false)
        }
    }

    //////////////////////////////////////////////////////////////// 
    // Row Tests
    //
    test("Row.getCell_") {
        val row = ExcelerWorkbook.create().sheet_(testSheet).row_(100)
        
        row.getCell_(100) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }

        val cell = row.cell_(100)

        row.getCell_(100) match {
            case Some(c) => assert(true)
            case None => assert(false)
        }
    }

    //////////////////////////////////////////////////////////////// 
    // Cell Tests
    //
    test("Cell.getValue_") {

        val workbook = ExcelerWorkbook.create()
        (
            for { 
                sheet <- workbook.createSheet_(testSheet)
                row = sheet.createRow(0)
                cell = row.createCell(0)
            } yield cell
        ) match { 
            case Success(cell) => {
                cell.setCellValue(testMessage)
            }
            case Failure(e) => println(e.getMessage())
        } 

        (
            for {
                sheet <- workbook.getSheet_(testSheet)
                row = sheet.getRow(0)
                cell = row.getCell(0)
            } yield cell
        ) match {
            case Some(cell) => {
                assert(cell.getValue_ == testMessage)
            }
            case None => assert(false)
        }
    }

    test("Cell.getUppper,Cell.upper_ etc") {
        val cell = ExcelerWorkbook.create().sheet_(testSheet).cell_(100, 100)

        assert(cell.getUpperCell_.isEmpty)
        assert(cell.getLowerCell_.isEmpty)
        assert(cell.getLeftCell_.isEmpty)
        assert(cell.getRightCell_.isEmpty)

        cell.upperCell_
        cell.lowerCell_
        cell.leftCell_
        cell.rightCell_

        assert(!cell.getUpperCell_.isEmpty)
        assert(!cell.getLowerCell_.isEmpty)
        assert(!cell.getLeftCell_.isEmpty)
        assert(!cell.getRightCell_.isEmpty)

    }

    test("Cell.getUpperStream_ etc") {
        val cell = ExcelerWorkbook.create().sheet_(testSheet).cell_(4, 4)

        assert(cell.getUpperStream_.take(10).toList.length == 1)
        assert(cell.getLowerStream_.take(10).toList.length == 1)
        assert(cell.getLeftStream_.take(10).toList.length == 1)
        assert(cell.getRightStream_.take(10).toList.length == 1)

        assert(cell.upperStream_.take(10).toList.length == 5)
        assert(cell.lowerStream_.take(10).toList.length == 10)
        assert(cell.leftStream_.take(10).toList.length == 5)
        assert(cell.rightStream_.take(10).toList.length == 10)

        assert(cell.getUpperStream_.take(10).toList.length == 5)
        assert(cell.getLowerStream_.take(10).toList.length == 10)
        assert(cell.getLeftStream_.take(10).toList.length == 5)
        assert(cell.getRightStream_.take(10).toList.length == 10)
    }

    test("Cell.hasBorderBottom_") {
        val workbook = ExcelerWorkbook.create()
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
        val workbook = ExcelerWorkbook.create()
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
        val workbook = ExcelerWorkbook.create()
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
        val workbook = ExcelerWorkbook.create()
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
        val workbook = ExcelerWorkbook.create()
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

}
