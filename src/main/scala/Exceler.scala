import org.apache.poi.ss.usermodel._
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import java.io._

import org.apache.commons.cli.CommandLine
import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.Options
import org.apache.commons.cli.ParseException

object Exceler {

    def parseCommandLine(args:Array[String]) : CommandLine = {
        val parser = new DefaultParser()
        val options = new Options()
        parser.parse(options, args)
    }

    def excel(filename:String) : Unit  = {
        val workbook = WorkbookFactory.create(new File(filename))
        val sheet = workbook.getSheet("test")
        val row = sheet.getRow(0)
        val cell = row.getCell(0)
        val value = cell.getCellType match {
            case Cell.CELL_TYPE_BLANK => cell.getStringCellValue
            case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue
            case Cell.CELL_TYPE_ERROR => cell.getErrorCellValue
            case Cell.CELL_TYPE_FORMULA => cell.getCellFormula
            case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue
            case Cell.CELL_TYPE_STRING => cell.getStringCellValue
            }
        println(value)
    }

    def main(args:Array[String]) : Unit  = {
        allCatch withTry { parseCommandLine(args) } match {
            case Success(cl) => excel(cl.getArgs()(0))
            case Failure(e) => println(e.getMessage())
        }
    }
}
