/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

object FileOp {

    def createDirectories(filename:String) : Path = {
        val dir = Paths.get(filename).getParent()
        if (dir != null)
            Files.createDirectories(dir)
        else
            null
    }

    def exists(filename:String) : Boolean = {
        Files.exists(Paths.get(filename))
    }
}

object ExcelImplicits {

    implicit class WorkbookImplicit(workbook:Workbook)  {

        def saveAs_(filename:String): Unit =  {
            FileOp.createDirectories(filename)
            val out = new FileOutputStream(filename)
            workbook.write(out)
            out.close()
        }

        def getSheet_(name:String):Option[Sheet] = Option(workbook.getSheet(name))

        def createSheet_(name:String):Try[Sheet] = Try(workbook.createSheet(name)) 

        def sheet_(name:String):Sheet = {
            this.getSheet_(name) match {
                case Some(s) => s
                case None => this.createSheet_(name) match {
                    case Success(s) => s
                    case Failure(e) => throw e
                }
            }
        }
    }

    implicit class CellImplicit(cell:Cell) {

        def getValue_():Any = {
            cell.getCellType match {
                case Cell.CELL_TYPE_BLANK => cell.getStringCellValue
                case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue
                case Cell.CELL_TYPE_ERROR => cell.getErrorCellValue
                case Cell.CELL_TYPE_FORMULA => cell.getCellFormula
                case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue
                case Cell.CELL_TYPE_STRING => cell.getStringCellValue
            }
        }
    }
}

object ExcelerWorkbook {

    def open(filename:String): Workbook =  {
        WorkbookFactory.create(new File(filename))
    }

    def create(): Workbook = new XSSFWorkbook()
}
