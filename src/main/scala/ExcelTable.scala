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


trait ExcelTableFunction extends TableFunction[ExcelRectangle] {
  val tableFunction = new FunctionImpl

  class FunctionImpl extends Function {
    override def getCross(row:ExcelRectangle, col:ExcelRectangle) = 
      ExcelRectangle(row.sheet, row.top, col.left, row.bottom, col.right)
  
    override def getHeadRow(rect:ExcelRectangle):(
      ExcelRectangle, Option[ExcelRectangle]) = {
      (for {
        rownum <- (rect.top until rect.bottom).toStream
        cell <- rect.sheet.getCellOption(rownum, rect.left)
        if cell.hasBorderBottom
      } yield rownum).headOption match {
        case Some(num) => (
          ExcelRectangle(rect.sheet, rect.top,
            rect.left, num, rect.right),
          Some(ExcelRectangle(rect.sheet, num + 1,
            rect.left, rect.bottom, rect.right)))
        case _ => (rect, None)
      }
    }
  
    override def getHeadCol(rect:ExcelRectangle):(
        ExcelRectangle,Option[ExcelRectangle]) = {
      (for {
        colnum <- (rect.left until rect.right).toStream
        cell <- rect.sheet.getCellOption(rect.top, colnum)
        if cell.hasBorderRight
      } yield colnum).headOption match {
        case Some(num) => (
          ExcelRectangle(rect.sheet, rect.top,
            rect.left, rect.bottom, num),
          Some(ExcelRectangle(rect.sheet, rect.top,
            num + 1, rect.bottom, rect.right)))
        case _ => (rect, None)
      }
    }
  
    override def getValue(rect:ExcelRectangle):Option[String] =
      (for {
        colnum <- (rect.left to rect.right).toStream
        rownum <- (rect.top to rect.bottom).toStream
        value <- rect.sheet.cell(rownum, colnum)
              .getValueString.map(_.trim)
      } yield value).headOption
  
    override def mergeRect(rectL:List[ExcelRectangle]):ExcelRectangle = {
      val head = rectL.head
      val last = rectL.last
      ExcelRectangle(
        head.sheet, head.top, head.left, last.bottom, last.right)
    }
  }
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
