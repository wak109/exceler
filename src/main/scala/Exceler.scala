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

class TableQueryImpl(
        val sheet:Sheet,
        val top:Int,
        val left:Int,
        val bottom:Int,
        val right:Int
    )
    extends ExcelRectangle
    with ExcelTableFunction
    with ExcelTableQueryFunction
    with StackedTableQuery[ExcelRectangle]
    with RectangleLineDraw 


object TableQueryImpl {
    implicit object factory extends ExcelFactory[TableQueryImpl] {
        override def create(
            sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int) =
                new TableQueryImpl(sheet,top,left,bottom,right)
    }
}

object Exceler extends ExcelTableFunction {
/*
    def convertExcelTableToXML(filename:String):Try[Unit] = {
        Try {
            val file = new File(filename)
            val workbook = WorkbookFactory.create(file, null ,true) 
            for {
                sheet <- workbook.sheetIterator.asScala
                rect <- sheet.getRectangleList[TableQueryImpl]
                row <- rect.getRowList(rect)
                cell <- row.getColumnList(row)
            } {
                println(cell)
            }
        }
    }
*/
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
                tableMap = sheet.getTableMap[TableQueryImpl]
                table <- tableMap.get(tablename)
                row <- table.query(
                    rowKeys.split(",").toList.map(isSameStr),
                    colKeys.split(",").toList.map(isSameStr))
                cell <- row
                value <- tableFunction.getValue(cell)
            } {
                println(value)
            }
        }
    }
}
