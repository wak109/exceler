/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */

package exceler.excel

import scala.collection._
import scala.language.implicitConversions
import scala.reflect.ClassTag
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

abstract trait TableLib[T] {
  def getRows(rect:T):List[T]
  def getColumns(rect:T):List[T]
  def getCross(row:T,column:T):T
  def merge(rectList:List[T]):T
}


object ExcelTableBorder {

  def reGroup(group:Map[(Int,Int),List[(Int,Int,Int)]]) = {
    for {
      key <- group.keys
    } yield (key, group.filterKeys(
          (x)=>((key._1 >= x._1) && (key._2 <= x._2)))
          .values.flatten.map(_._1).toList.sorted)
  }

  def getRows(rect:ExcelRectangle):List[(Int,Int)] = {
    val groups = reGroup(getHorizontalLines(rect).groupBy(
          (t)=>(t._2, t._3)))
    (((rect.top-1)::groups.map(_._2).maxBy(_.length)):+rect.bottom)
      .pairwise.map((tpl)=>(tpl._1+1,tpl._2))
  }

  def getColumns(rect:ExcelRectangle):List[(Int,Int)] = {
    val groups = reGroup(getVerticalLines(rect).groupBy(
          (t)=>(t._2, t._3)))
    (((rect.left-1)::groups.map(_._2).maxBy(_.length)):+rect.right)
      .pairwise.map((tpl)=>(tpl._1+1,tpl._2))
  }

  def getHorizontalLines(rect:ExcelRectangle):List[(Int,Int,Int)] =
    for {
      rownum <- (rect.top until rect.bottom).toList
      line <- rect.row(rownum).toList.blockingBy(_.hasBorderBottom)
    } yield (rownum, line.head.getColumnIndex, line.last.getColumnIndex)

  def getVerticalLines(rect:ExcelRectangle):List[(Int,Int,Int)] =
    for {
      colnum <- (rect.left until rect.right).toList
      line <- rect.column(colnum).toList.blockingBy(_.hasBorderRight)
    } yield (colnum, line.head.getRowIndex, line.last.getRowIndex)
}


trait ExcelTableLibImpl extends TableLib[ExcelRectangle] {

  override def getRows(rect:ExcelRectangle):List[ExcelRectangle] = 
    ExcelTableBorder.getRows(rect).map(tpl => ExcelRectangle(
      rect.sheet, tpl._1, rect.right, tpl._2, rect.right))

  override def getColumns(rect:ExcelRectangle):List[ExcelRectangle] =
    ExcelTableBorder.getColumns(rect).map(tpl => ExcelRectangle(
      rect.sheet, rect.top, tpl._1, rect.bottom, tpl._2))

  override def getCross(
      row:ExcelRectangle, column:ExcelRectangle):ExcelRectangle =
    ExcelRectangle(row.sheet,
      row.top, column.right, row.bottom, column.left)

  override def merge(rectList:List[ExcelRectangle]):ExcelRectangle =
    ExcelRectangle(
        rectList.head.sheet,
        rectList.head.top,
        rectList.head.left,
        rectList.last.bottom,
        rectList.last.right)
}


class Rect2(
    val sheet:Sheet, val top:Int, val left:Int,
    val bottom:Int, val right:Int)
      extends ExcelRectangle with TableQueryTraitImpl {

  def getValue():Option[String] =
    this.getValue(this.sheet, this.top, this.left, this.bottom, this.right)
}

object Rect2 {
  implicit object rectTableLib extends TableLib[Rect2] {
    val func = new ExcelTableLibImpl {}
    def conv(rect:ExcelRectangle) = new Rect2(rect.sheet,rect.top,rect.left,rect.bottom,rect.right)

    override def getRows(rect:Rect2):List[Rect2] = func.getRows(rect).map(conv)
    override def getColumns(rect:Rect2):List[Rect2] = func.getColumns(rect).map(conv)
    override def getCross(row:Rect2,column:Rect2):Rect2 = conv(func.getCross(row,column))
    override def merge(rectList:List[Rect2]):Rect2 = conv(func.merge(rectList))
  }
}

class TableQuery2[
    T<:{def getValue():Option[String]}:TableLib:ClassTag](val rect:T)
{
  val tlib = implicitly[TableLib[T]]

  val rows = tlib.getRows(rect)
  val columns = tlib.getColumns(rect)
  val cells = List.tabulate(rows.length, columns.length)(
    (row, col)=>tlib.getCross(rows(row), columns(col)))
}
