/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._
import ExcelLib.Table.ImplicitConversions._

object Exceler {

    def convertExcelTableToXML(filename:String):Try[Unit] = {
        Try {
            val file = new File(filename)
            val workbook = WorkbookFactory.create(file, null ,true) 
            for {
                sheet <- workbook.sheetIterator.asScala
                rect <- sheet.getRectangleList[ExcelTable]
                row <- rect.getRowList[ExcelTable]
                cell <- row.getColumnList[ExcelTable]
            } {
                println(cell)
            }
        }
    }

    def readExcelTable(
        filename:String,
        sheetname:String,
        tablename:String,
        rowKeys:String,
        colKeys:String):Try[Unit] = {
        Try {
            def isSameStr(s:String): String => Boolean = {
                s match {
                    case "" => (x:String) => true
                    case _  => (x:String) => x == s
                }
            }
            val file = new File(filename)
            val workbook = WorkbookFactory.create(file, null ,true) 

            for {
                sheet <- workbook.getSheetOption(sheetname)
                tableMap = sheet.getTableMap
                table <- tableMap.get(tablename)
                row <- table.query(
                    rowKeys.split(",").toList.map(isSameStr),
                    colKeys.split(",").toList.map(isSameStr))
                cell <- row
                value <- TableFunctionImpl.getValue(cell)
            } {
                println(value)
            }
        }
    }
}
