/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import CommonLib.ImplicitConversions._
import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._


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


trait TableQueryTraitImpl
    extends TableQueryTrait[ExcelRectangle]
    with ExcelTableTraitImpl {

  def getValue(sheet:Sheet,row:Int,col:Int):Option[String] = 
    this.getValue(sheet, row, col, row, col)

  override def getValue(rect:ExcelRectangle):Option[String] = 
    this.getValue(rect.sheet,rect.top,rect.left,rect.bottom,rect.right)
    
  def getValue(
    sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int):Option[String] = {
    val topLeft = this.getTopLeft(
      sheet,top,left).getOrElse((top, left))
    val bottomRight = this.getBottomRight(
      sheet,bottom,right).getOrElse((bottom, right))

    (for {
      col <- (topLeft._2 to bottomRight._2).toStream
      row <- (topLeft._1 to bottomRight._1).toStream
      value <- sheet.cell(row, col).getValueString.map(_.trim)
    } yield value).headOption
  }

  def getTopLeft(sheet:Sheet,row:Int,col:Int):Option[(Int,Int)] =
    (for {
      cell <- sheet.cell(row,col).getLeftStream if cell.hasBorderLeft
      cell <- cell.getUpperStream if cell.hasBorderTop
    } yield (cell.getRowIndex, cell.getColumnIndex)).headOption

  def getBottomRight(sheet:Sheet,row:Int,col:Int):Option[(Int,Int)] = 
    (for {
      cell <- sheet.cell(row,col).getRightStream if cell.hasBorderRight
      cell <- cell.getLowerStream if cell.hasBorderBottom
    } yield (cell.getRowIndex, cell.getColumnIndex)).headOption
}
