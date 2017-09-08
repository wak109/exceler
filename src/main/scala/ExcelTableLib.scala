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

  /**
   */ 
  def getHorizontalLines(rect:ExcelRectangle):List[(Int,Int,Int)] =
    for {
      rownum <- (rect.top until rect.bottom).toList
      line <- rect.row(rownum).toList.blockingBy(_.hasBorderBottom)
    } yield (rownum, line.head.getColumnIndex, line.last.getColumnIndex)

  /**
   */
  def getVerticalLines(rect:ExcelRectangle):List[(Int,Int,Int)] =
    for {
      colnum <- (rect.left until rect.right).toList
      line <- rect.column(colnum).toList.blockingBy(_.hasBorderRight)
    } yield (colnum, line.head.getRowIndex, line.last.getRowIndex)
}



abstract trait TableTrait[T] {
  def getRows(rect:T):List[T]
  def getColumns(rect:T):List[T]
  def getCross(row:T,column:T):T
  def merge(rectList:List[T]):T
}


trait ExcelTableTraitImpl extends TableTrait[ExcelRectangle] {

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


class Rect(
    val sheet:Sheet, val top:Int, val left:Int,
    val bottom:Int, val right:Int)
      extends ExcelRectangle with TableQueryTraitImpl {

  def getValue():Option[String] =
    this.getValue(this.sheet, this.top, this.left, this.bottom, this.right)
}

object Rect {
  implicit object rectTableTrait extends TableTrait[Rect] {
    val func = new ExcelTableTraitImpl {}
    def conv(rect:ExcelRectangle) = new Rect(rect.sheet,rect.top,rect.left,rect.bottom,rect.right)

    override def getRows(rect:Rect):List[Rect] = func.getRows(rect).map(conv)
    override def getColumns(rect:Rect):List[Rect] = func.getColumns(rect).map(conv)
    override def getCross(row:Rect,column:Rect):Rect = conv(func.getCross(row,column))
    override def merge(rectList:List[Rect]):Rect = conv(func.merge(rectList))
  }
}
