/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import java.nio.file.{Paths, Files}

import ExcelImplicits._

class ExcelerSuite extends FunSuite with BeforeAndAfterEach {
  
    val testFile = "testFile.xlsx"
    val testSheet = "testSheet"
    val testMessage = "Hello, world"

    override def beforeEach() {
        Files.deleteIfExists(Paths.get(testFile))
    }
    
    override def afterEach() {
        Files.deleteIfExists(Paths.get(testFile))
    }
    
    test("Workbook saveAs_") {
        val workbook = ExcelerWorkbook.create()
        workbook.saveAs_(testFile)
        assert(Files.exists(Paths.get(testFile)))
    }

    test("Create Excel Sheet") {
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

    test("WorkbookImplicit sheet_") {
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

    test("SheetImplicit") {
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

    test("SheetImplicit getCell_") {
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

    test("RowImplicit") {
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

    test("Create Set Cell Value") {
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
        workbook.saveAs_(testFile)
        assert(Files.exists(Paths.get(testFile)))


        val workbook2 = ExcelerWorkbook.open(testFile)
        (
            for {
                sheet <- workbook2.getSheet_(testSheet)
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

    test("CellImplicit uppper lower left right") {
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

    test("CellImplicit Stream uppper lower left right ") {
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
}
