/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import org.scalatest._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.poi.POIXMLDocumentPart

import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib._

class ExcelLibSuite extends FunSuite with BeforeAndAfterEach {
  
    val testWorkbook1 = "test1.xlsx"

    val testSheet = "test"
    val testMessage = "Hello, world!!"

    override def beforeEach() {
    }
    
    override def afterEach() {
    }
    
    //////////////////////////////////////////////////////////////// 
    // Workbook Tests
    //
    test("Workbook.saveAs") {
        val testFile = "Test_Workbook_saveAs.xlsx"

        val workbook = new XSSFWorkbook()
        workbook.saveAs(testFile)
        assert(Files.exists(Paths.get(testFile)))

        Files.deleteIfExists(Paths.get(testFile))
    }

    test("Workbook.getSheet_") {
        val workbook = new XSSFWorkbook()

        workbook.getSheetOption(testSheet) match {
            case Some(_) => assert(false)
            case None => assert(true)
        }

        workbook.sheet(testSheet)

        workbook.getSheetOption(testSheet) match {
            case Some(_) => assert(true)
            case None => assert(false)
        }
    }

    test("Workbook.sheet") {
        val workbook = new XSSFWorkbook()

        Try(workbook.sheet(testSheet)) match {
            case Success(s) => assert(s.getSheetName() == testSheet)
            case Failure(e) => assert(false)
        }

        Try(workbook.sheet(testSheet)) match {
            case Success(s) => assert(s.getSheetName() == testSheet)
            case Failure(e) => assert(false)
        }
        assert(workbook.sheet(testSheet) == workbook.getSheet(testSheet))
    }

    test("Workbook.findCellStyle") {
        val workbook = new XSSFWorkbook()

        val cellStyle = workbook.sheet(testSheet).cell(5, 5).getCellStyle
        assert(workbook.findCellStyle(cellStyle).nonEmpty)

        val tuple = cellStyle.toTuple
        assert(workbook.findCellStyle(tuple).nonEmpty)

        assert(workbook.findCellStyle(
            (tuple._1, tuple._2, tuple._3, BorderStyle.THIN, tuple._5,
            tuple._6, tuple._7, tuple._8, tuple._9, tuple._10)
            ).isEmpty)

        var newStyle = workbook.createCellStyle
        newStyle.cloneStyleFrom(cellStyle)
        newStyle.setBorderRight(BorderStyle.THIN)

        assert(workbook.findCellStyle(newStyle).nonEmpty)

        assert(workbook.findCellStyle(
            (tuple._1, tuple._2, tuple._3, BorderStyle.THIN, tuple._5,
            tuple._6, tuple._7, tuple._8, tuple._9, tuple._10)
            ).nonEmpty)
    }


    //////////////////////////////////////////////////////////////// 
    // Sheet Tests
    //
    test("Sheet.getRow_,Sheet.row_") {
        val workbook = new XSSFWorkbook()
        val sheet = workbook.sheet(testSheet)

        sheet.getRowOption(0) match {
            case Some(r) => assert(false)
            case None => assert(true)
        }

        val row = sheet.row(0)

        sheet.getRowOption(0) match {
            case Some(r) => assert(true)
            case None => assert(false)
        }
    }

    test("Sheet.getCell_") {
        val sheet = (new XSSFWorkbook()).sheet(testSheet)

        sheet.getCellOption(0, 0) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }

        val cell = sheet.cell(0, 0)

        sheet.getCellOption(0, 0) match {
            case Some(c) => assert(true)
            case None => assert(false)
        }
    }

    test("Sheet.getDrawingPatriarch") {
        val sheet = (new XSSFWorkbook()).sheet(testSheet)

        sheet.getDrawingPatriarchOption match {
            case Some(c) => assert(false)
            case None => assert(true)
        }

        val drawing = sheet.drawingPatriarch

        println(drawing.getClass)


        sheet.getDrawingPatriarchOption match {
            case Some(c) => assert(true)
            case None => assert(false)
        }
    }

    //////////////////////////////////////////////////////////////// 
    // Row Tests
    //
    test("Row.getCell_") {
        val row = (new XSSFWorkbook()).sheet(testSheet).row(100)
        
        row.getCellOption(100) match {
            case Some(c) => assert(false)
            case None => assert(true)
        }

        val cell = row.cell(100)

        row.getCellOption(100) match {
            case Some(c) => assert(true)
            case None => assert(false)
        }
    }

    //////////////////////////////////////////////////////////////// 
    // Cell Tests
    //
    test("Cell.getValue_") {

        val workbook = new XSSFWorkbook()
        val sheet = workbook.sheet(testSheet)
        val row = sheet.createRow(0)
        val cell = row.createCell(0)

        cell.setCellValue(testMessage)

        (
            for {
                sheet <- workbook.getSheetOption(testSheet)
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
        val cell = (new XSSFWorkbook()).sheet(testSheet).cell(100, 100)

        assert(cell.getUpperCell.isEmpty)
        assert(cell.getLowerCell.isEmpty)
        assert(cell.getLeftCell.isEmpty)
        assert(cell.getRightCell.isEmpty)

        cell.upperCell
        cell.lowerCell
        cell.leftCell
        cell.rightCell

        assert(!cell.getUpperCell.isEmpty)
        assert(!cell.getLowerCell.isEmpty)
        assert(!cell.getLeftCell.isEmpty)
        assert(!cell.getRightCell.isEmpty)

    }

    test("Cell.getUpperStream etc") {
        val cell = (new XSSFWorkbook()).sheet(testSheet).cell(4, 4)

        assert(cell.getUpperStream.take(10).toList.length == 5)
        assert(cell.getLowerStream.take(10).toList.length == 10)
        assert(cell.getLeftStream.take(10).toList.length == 5)
        assert(cell.getRightStream.take(10).toList.length == 10)

        assert(cell.upperStream.take(10).toList.length == 5)
        assert(cell.lowerStream.take(10).toList.length == 10)
        assert(cell.leftStream.take(10).toList.length == 5)
        assert(cell.rightStream.take(10).toList.length == 10)

        assert(cell.getUpperStream.take(10).toList.length == 5)
        assert(cell.getLowerStream.take(10).toList.length == 10)
        assert(cell.getLeftStream.take(10).toList.length == 5)
        assert(cell.getRightStream.take(10).toList.length == 10)
    }
    
    test("Shapes") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val relations = workbook.getSheet("shapes").asInstanceOf[POIXMLDocumentPart].getRelations.asScala
        val drawing = (
            for {
                xfd <- relations
                if xfd.isInstanceOf[XSSFDrawing]
            } yield xfd.asInstanceOf[XSSFDrawing]
        ).headOption 
        drawing match {
            case Some(xd) => assert(xd.getShapes.asScala.length == 8)
            case None => assert(false)
        }
    }

    test("getShapes") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheet("shapes")
        assert(sheet.getXSSFShapes.length == 8)
    }

    test("XSSFShapeExt") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheet("shapes")
        for (shape <- sheet.getXSSFShapes) {
           println(shape.toXmlObject)
        }
    }
}
