/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
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

import CommonLib._


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

      val result = for {
        book <- ExcelerBook.getBook(filename).toSeq
        sheet <- book.getSheet(sheetname).toSeq
        table <- sheet.getTable(tablename).toSeq
        row <- table.query(
          rowKeys.split(",").toList.map(isSameStr),
          colKeys.split(",").toList.map(isSameStr))
        cell <- row
        value <- tableFunction.getValue(cell)
      } yield value

      println(result)
    }
  }
}


class ExcelerSheet(_sheet:Sheet) {
  val sheet = _sheet
  lazy private val tableMap = sheet.getTableMap[TableQueryImpl]

  def getTable(tableName:String) = tableMap.get(tableName)
}


class ExcelerBook(_workbook:Workbook) {

  val workbook = _workbook
  lazy private val sheetMap = (for {
      i <- Range(0, workbook.getNumberOfSheets)
    } yield (workbook.getSheetName(i), new ExcelerSheet(
        workbook.getSheetAt(i)))).toMap

  def getSheet(sheetName:String) = sheetMap.get(sheetName)
}

object ExcelerBook {

  lazy private val bookMap = 
    getListOfFiles(Config().excelDir)
      .filter(_.canRead)
      .filter(_.getName.endsWith(".xlsx"))
      .map((f:File)=>(f.getName, new ExcelerBook(
        WorkbookFactory.create(f, null ,true))))
      .toMap

  def getBook(bookName:String) = bookMap.get(bookName)
}

