import org.apache.poi.ss.usermodel._
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import java.io._

import org.apache.commons.cli.CommandLine
import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.HelpFormatter
import org.apache.commons.cli.{Option => CmdOption}
import org.apache.commons.cli.{Options => CmdOptions}
import org.apache.commons.cli.ParseException

object Exceler {

    val description = """Scala CLI template"""
    val access = """https://github.com/wak109/"""

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

    def checkOptions(cl:CommandLine) : Unit = {
        if (cl.hasOption('h')) {
            printUsage()
        } else if (cl.getArgs().length == 0) {
            printUsage()
        } else {
            excel(cl.getArgs()(0))
        }
    }

    def stripClassName(clsname:String):String = {
        val Pattern = """^(.*)\$$""".r
        clsname match {
          case Pattern(m) => m
          case x => x
        }
    }

    def printUsage() : Unit = {
        val formatter = new HelpFormatter()

        formatter.printHelp(
            stripClassName(this.getClass.getCanonicalName),
            this.description,
            makeOptions(),
            this.access,
            true)
    }

    def parseCommandLine(args:Array[String]) : CommandLine = {
        val parser = new DefaultParser()

        return parser.parse(makeOptions(), args)
    }

    def makeOptions() : CmdOptions = {
        val options = new CmdOptions()
        options.addOption(CmdOption.builder("h").desc("Show help").build())

        return options
    }

    def main(args:Array[String]) : Unit  = {
        allCatch withTry { parseCommandLine(args) } match {
            case Success(cl) => checkOptions(cl)
            case Failure(e) => println(e.getMessage()); printUsage()
        }
    }
}
