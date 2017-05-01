/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest._

import java.nio.file.{Paths, Files}

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
    
    test("Create Excel Workbook") {
        val workbook = ExcelerWorkbook.create()
        ExcelerWorkbook.saveAs(workbook, testFile)
        assert(Files.exists(Paths.get(testFile)))
    }

    test("Create Excel Sheet") {
        val workbook = ExcelerWorkbook.create()

        ExcelerWorkbook.getSheet(workbook, testSheet) match {
            case Some(_) => assert(false)
            case None => assert(true)
        }

        ExcelerWorkbook.createSheet(workbook, testSheet) match {
            case Success(_) => assert(true)
            case Failure(e) => assert(false)
        }

        ExcelerWorkbook.getSheet(workbook, testSheet) match {
            case Some(_) => assert(true)
            case None => assert(false)
        }

        /* Can't create sheets of same name */
        ExcelerWorkbook.createSheet(workbook, testSheet) match {
            case Success(_) => assert(false)
            case Failure(e) => assert(true)
        }
    }

    test("Create Set Cell Value") {
        val workbook = ExcelerWorkbook.create()
        (
            for { 
                sheet <- ExcelerWorkbook.createSheet(workbook, testSheet)
                row = sheet.createRow(0)
                cell = row.createCell(0)
            } yield cell
        ) match { 
            case Success(cell) => {
                cell.setCellValue(testMessage)
            }
            case Failure(e) => println(e.getMessage())
        } 
        ExcelerWorkbook.saveAs(workbook, testFile)
        assert(Files.exists(Paths.get(testFile)))
    }
}
