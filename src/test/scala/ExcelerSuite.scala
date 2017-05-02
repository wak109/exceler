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
}
