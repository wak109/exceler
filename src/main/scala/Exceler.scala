/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.excel

import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._
import excellib.Table.ImplicitConversions._

import exceler._
import exceler.common._
import exceler.table._

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


object TableQueryImpl extends ExcelTableFunction {
  implicit object factory extends ExcelFactory[TableQueryImpl] {
    override def create(
      sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int) =
        new TableQueryImpl(sheet,top,left,bottom,right)
  }

  def queryExcelTable(
    filename:String,
    sheetname:String,
    tablename:String,
    rowKeys:String,
    colKeys:String) = {
    def isSameStr(s:String): String => Boolean = {
      s match {
        case "" => (x:String) => true
        case _  => (x:String) => x == s
      }
    }

    Try {
      for {
        book <- ExcelerBook.getBook(filename).toSeq
        sheet <- book.getSheetOption(sheetname).toSeq
        table <- sheet.getTableMap[TableQueryImpl].get(tablename).toSeq
        row <- table.query(
          rowKeys.split(",").toList.map(isSameStr),
          colKeys.split(",").toList.map(isSameStr))
        cell <- row
        value <- tableFunction.getValue(cell)
      } yield value
    }
  }
}


object ExcelerBook {

  lazy private val bookMap = 
    getListOfFiles(Config().excelDir)
      .filter(_.canRead)
      .filter(_.getName.endsWith(".xlsx"))
      .map((f:File)=>(f.getName, WorkbookFactory.create(f, null ,true)))
      .toMap

  def getBook(bookName:String) = bookMap.get(bookName)
}

