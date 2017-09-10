/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.excel

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import exceler.common._
import exceler.table._
import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._


class ExcelTableQuery (val sheet:Sheet, val top:Int, val left:Int,
  val bottom:Int, val right:Int)
    extends TableQuery[ExcelRectangle]
    with ExcelRectangle
    with ExcelTableQueryFunction {
  override val tableQueryFunction = new QueryFunctionImpl{}
}


trait ExcelTableQueryFunction
  extends ExcelTableFunction 
  with TableQueryFunction[ExcelRectangle] {

  override val tableQueryFunction = new QueryFunctionImpl{}

  trait QueryFunctionImpl extends QueryFunction {

    override def create(rect:ExcelRectangle) =
      new ExcelTableQuery(rect.sheet, rect.top,
          rect.left, rect.bottom, rect.right)
  }
}


trait TableQueryTraitImpl {

  def getValue(rect:ExcelRectangle):Option[String] = 
    this.getValue(rect.sheet,rect.top,rect.left,rect.bottom,rect.right)
    
  def getValue(sheet:Sheet,row:Int,col:Int):Option[String] = 
    this.getValue(sheet, row, col, row, col)

  def getValue(
    sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int):Option[String] = {

    val topLeft = sheet.cell(top,left).getTopLeftCorner.fold(
      (top, left))(cell => (cell.getRowIndex, cell.getColumnIndex))
    val bottomRight = sheet.cell(bottom,right).getBottomRightCorner.fold(
      (bottom, right))(cell => (cell.getRowIndex, cell.getColumnIndex))

    (for {
      col <- (topLeft._2 to bottomRight._2).toStream
      row <- (topLeft._1 to bottomRight._1).toStream
      value <- sheet.cell(row, col).getValueString.map(_.trim)
    } yield value).headOption
  }
}
