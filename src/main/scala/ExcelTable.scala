/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
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

package ExcelLib.Table {
    trait ImplicitConversions {
        implicit class ToExcelTableSheetConversion(val sheet:Sheet)
                extends ExcelTableSheetConversion
    }
    object ImplicitConversions extends ImplicitConversions
}

trait Factory[T] {
    type Base
    def create(rect:Base):T
}

trait ExcelFactory[T] extends Factory[T] {
    type Base = ExcelRectangle
    def create(sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int):T
    def create(rect:Base):T = this.create(
        rect.sheet, rect.top, rect.left, rect.bottom, rect.right)
}

trait ExcelRectangle {
    val sheet:Sheet
    val top:Int
    val left:Int
    val bottom:Int
    val right:Int
}


object ExcelRectangle {
    implicit val function = new ExcelTableFunction{}

    implicit def factory(rect:ExcelRectangle) = new TableQueryImpl(
        rect.sheet, rect.top, rect.left, rect.bottom, rect.right)

    def apply(sheet:Sheet, top:Int, left:Int, bottom:Int, right:Int) = {

        class Impl (
            val sheet:Sheet,
            val top:Int,
            val left:Int,
            val bottom:Int,
            val right:Int
        ) extends ExcelRectangle

        new Impl(sheet, top, left, bottom, right)
    }
}

trait ExcelTableFunction extends TableFunction[ExcelRectangle] {

    override def getCross(row:ExcelRectangle, col:ExcelRectangle) = {
        ExcelRectangle(row.sheet, row.top, col.left, row.bottom, col.right)
    }

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

    override def getTableName(rect:ExcelRectangle)
            : (Option[String], ExcelRectangle) = {
        // Table name outside of Rectangle
        (rect.top match {
            case 0 => (None, rect)
            case _ => (rect.sheet.cell(rect.top - 1, rect.left)
                    .getValueString, rect)
        }) match {
            case (Some(name), _) => (Some(name), rect)
            case (None, _) => {
                // Table name at the top of Rectangle
                val (rowHead, rowTail) = this.getHeadRow(rect)
                val (rowHeadLeft, rowHeadRight) = this.getHeadCol(rowHead)
                rowHeadRight match {
                    case Some(_) => (None, rect)
                    case None => (this.getValue(rowHeadLeft),
                            rowTail) match {
                        case (name, Some(tail)) => (name, tail)
                        case (name, None) => (None, rect)
                    }
                }
            }
        }
    }

    override def mergeRect(rectL:List[ExcelRectangle]):ExcelRectangle = {
        val head = rectL.head
        val last = rectL.last
        ExcelRectangle(
            head.sheet, head.top, head.left, last.bottom, last.right)
    }
}


class TableQueryImpl(
        val sheet:Sheet,
        val top:Int,
        val left:Int,
        val bottom:Int,
        val right:Int
    )
    extends TableComponent[ExcelRectangle]
    with ExcelRectangle
    with Table[ExcelRectangle]
    with StackedTableQuery[ExcelRectangle]
    with RectangleLineDraw 


object TableQueryImpl {

    implicit object factory extends ExcelFactory[TableQueryImpl] {
        override def create(
            sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int) =
                new TableQueryImpl(sheet,top,left,bottom,right)
    }
}


trait RectangleLineDraw {
    rect:ExcelRectangle =>

    def drawOuterBorderTop(borderStyle:BorderStyle):Unit = {
        for (colnum <- (rect.left to rect.right).toList)
            rect.sheet.cell(rect.top, colnum).setBorderTop(borderStyle)
    }

    def drawOuterBorderLeft(borderStyle:BorderStyle):Unit = {
        for (rownum <- (rect.top to rect.bottom).toList)
            rect.sheet.cell(rownum, rect.left).setBorderLeft(borderStyle)
    }

    def drawOuterBorderBottom(borderStyle:BorderStyle):Unit = {
        for (colnum <- (rect.left to rect.right).toList)
            rect.sheet.cell(rect.bottom, colnum)
                .setBorderBottom(borderStyle)
    }

    def drawOuterBorderRight(borderStyle:BorderStyle):Unit = {
        for (rownum <- (rect.top to rect.bottom).toList)
            rect.sheet.cell(rownum, rect.right)
                .setBorderRight(borderStyle)
    }

    def drawOuterBorder(borderStyle:BorderStyle):Unit = {
        drawOuterBorderTop(borderStyle)
        drawOuterBorderLeft(borderStyle)
        drawOuterBorderBottom(borderStyle)
        drawOuterBorderRight(borderStyle)
    }

    def drawHorizontalLine(num:Int, borderStyle:BorderStyle):Unit = {
        if (num == 0) {
            for (colnum <- (rect.left to rect.right).toList)
                rect.sheet.cell(0, colnum).setBorderTop(borderStyle)
        }
        else if (0 < num && num <= rect.bottom - rect.top) {
            for (colnum <- (rect.left to rect.right).toList)
                rect.sheet.cell(rect.top + num - 1, colnum)
                    .setBorderBottom(borderStyle)
        }
    }

    def drawVerticalLine(num:Int, borderStyle:BorderStyle):Unit = {
        if (num == 0) {
            for (rownum <- (rect.top to rect.bottom).toList)
                rect.sheet.cell(rownum, 0).setBorderLeft(borderStyle)
        }
        else if (0 < num && num <= rect.right - rect.left) {
            for (rownum <- (rect.top to rect.bottom).toList)
                rect.sheet.cell(rownum, rect.left + num - 1)
                    .setBorderRight(borderStyle)
        }
    }
}

trait ExcelTableSheetConversion {
    val sheet:Sheet

    def getTableMap[T:ExcelFactory]():Map[String,T] = {
        val function = new ExcelTableFunction{}
        val factory = implicitly[ExcelFactory[T]]
        val tableList = sheet.getRectangleList

        tableList.map(t=>function.getTableName(t)).zipWithIndex.map(
            _ match {
                case ((Some(name), t), _) => (name, factory.create(t))
                case ((None, t), idx) => ("Table" + idx, factory.create(t))
            }).toMap
    }
}
