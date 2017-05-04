/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConversions._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib._
import ExcelTable._


object Exceler {

    def convertExcelTableToXML(filename:String):Unit = {
        val workbook = WorkbookFactory.create(new File(filename))
        for {
            sheet <- workbook.sheetIterator
            table <- getExcelTableList(sheet)
        } {
            println(table)
        }
    }
}
