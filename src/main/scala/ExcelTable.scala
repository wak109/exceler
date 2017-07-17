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

trait TableFunction[T] {
    def getCross(row:T, col:T):T
    def getHeadRow(rect:T):(T,Option[T])
    def getHeadCol(rect:T):(T,Option[T])
    def getValue(rect:T):Option[String]
    def getTableName(rect:T):(Option[String], T)
    def mergeRect(rectL:List[T]):T
}


trait Table[T] {
    rect:T =>
    val tableFunc:TableFunction[T]

    lazy val rowList:List[T] = this.getRowList(Some(rect))
    lazy val colList:List[T] = this.getColList(Some(rect))

    def getRowList(rect:Option[T]):List[T] = {
        rect match {
            case None => Nil
            case Some(rect) => {
                val (headRow, tailRow) = tableFunc.getHeadRow(rect)
                headRow :: getRowList(tailRow)
            }
        }
    }

    def getColList(rect:Option[T]):List[T] = {
        rect match {
            case None => Nil
            case Some(rect) => {
                val (headCol, tailCol) = tableFunc.getHeadCol(rect)
                headCol :: getColList(tailCol)
            }
        }
    }
}

trait TableQuery[T] extends Table[T] {
    rect:T =>
    val tableFunc:TableFunction[T]
    val createTableQuery:(T) => TableQuery[T]

    def query(
        rowpredList:List[(String) => Boolean],
        colpredList:List[(String) => Boolean]
    ):List[List[T]] = {
        for {
            row <- this.queryRow(rowpredList)
        } yield for {
            col <- this.queryColumn(colpredList)
        } yield tableFunc.getCross(row, col)
    }

    def query(
        rowpred:(String) => Boolean,
        colpred:(String) => Boolean
    ):List[List[T]] = {
        for {
            row <- this.queryRow(rowpred)
        } yield for {
            col <- this.queryColumn(colpred)
        } yield tableFunc.getCross(row, col)
    }

    def queryRow(
        predList:List[String => Boolean]
    ):List[T] = {
        predList match {
            case Nil => List[T](rect)
            case pred::predTail => for {
                row <- this.queryRow(pred)
                row <- createTableQuery(row).queryRow(predTail)
            } yield row
        }
    }

    def queryRow(
        pred:String => Boolean
    ):List[T] = {
        for {
            row <- this.rowList
            (colHead, colTail) = tableFunc.getHeadCol(row)
            if pred(tableFunc.getValue(colHead).getOrElse(""))
            col <- colTail
        } yield col
    }

    def queryColumn(
            predList:List[String => Boolean]
    ):List[T] = {
        predList match {
            case Nil => List[T](rect)
            case pred::predTail => for {
                col <- this.queryColumn(pred)
                col <- createTableQuery(col).queryColumn(predTail)
            } yield col
        }
    }

    def queryColumn(
        pred:String => Boolean
    ) :List[T] = {
        for {
            col <- this.colList
            (rowHead, rowTail) = tableFunc.getHeadRow(col)
            if pred(tableFunc.getValue(rowHead).getOrElse(""))
            row <- rowTail
        } yield row
    }
}

trait StackedTableQuery[T] extends TableQuery[T] {
    rect:T =>

    val tableFunc:TableFunction[T]
    val createTableQuery:(T) => TableQuery[T]
    val isSeparator:(T)=>Boolean = tableFunc.getHeadCol(_)._2 == None

    lazy val tableMap = this.rowList
        .splitBy(isSeparator)
        .pairingBy(x=>isSeparator(tableFunc.mergeRect(x)))
        .map(pair=>(tableFunc.getValue(
            tableFunc.mergeRect(pair._1)).getOrElse(""),
            tableFunc.mergeRect(pair._2)))
        .map(pair => (pair._1, createTableQuery(pair._2)))
        .toMap

    override def queryRow(predList:List[String => Boolean]):List[T] = {
        predList match {
            case Nil => List[T](rect)
            case pred::predTail => {
                (for {
                    (key, table) <- tableMap
                    if pred(key)
                    row <- table.queryRow(predTail) 
                } yield row).toList match {
                    case Nil => super.queryRow(predList)
                    case x => x
                }
            }
        }
    }
}

////////////////////////////////////////////////////////////////////////
// Impl
//

case class RectangleImpl( 
    val sheet:Sheet,
    val top:Int,
    val left:Int,
    val bottom:Int,
    val right:Int
) 

object TableFunctionImpl extends TableFunction[RectangleImpl] {

    override def getCross(row:RectangleImpl, col:RectangleImpl) = {
        new RectangleImpl(
                row.sheet, row.top, col.left, row.bottom, col.right)
    }

    override def getHeadRow(rect:RectangleImpl):(
        RectangleImpl, Option[RectangleImpl]) = {
        (for {
            rownum <- (rect.top until rect.bottom).toStream
            cell <- rect.sheet.getCellOption(rownum, rect.left)
            if cell.hasBorderBottom
        } yield rownum).headOption match {
            case Some(num) => (
                new RectangleImpl(rect.sheet, rect.top,
                    rect.left, num, rect.right),
                Some(new RectangleImpl(rect.sheet, num + 1,
                    rect.left, rect.bottom, rect.right)))
            case _ => (rect, None)
        }
    }

    override def getHeadCol(rect:RectangleImpl):(
            RectangleImpl,Option[RectangleImpl]) = {
        (for {
            colnum <- (rect.left until rect.right).toStream
            cell <- rect.sheet.getCellOption(rect.top, colnum)
            if cell.hasBorderRight
        } yield colnum).headOption match {
            case Some(num) => (
                new RectangleImpl(rect.sheet, rect.top,
                    rect.left, rect.bottom, num),
                Some(new RectangleImpl(rect.sheet, rect.top,
                    num + 1, rect.bottom, rect.right)))
            case _ => (rect, None)
        }
    }

    override def getValue(rect:RectangleImpl):Option[String] =
        (for {
            colnum <- (rect.left to rect.right).toStream
            rownum <- (rect.top to rect.bottom).toStream
            value <- rect.sheet.cell(rownum, colnum)
                        .getValueString.map(_.trim)
        } yield value).headOption

    override def getTableName(rect:RectangleImpl)
            : (Option[String], RectangleImpl) = {
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

    override def mergeRect(rectL:List[RectangleImpl]):RectangleImpl = {
        val head = rectL.head
        val last = rectL.last
        new RectangleImpl(
            head.sheet, head.top, head.left, last.bottom, last.right)
    }
}


class TableQueryImpl(
        sheet:Sheet,
        top:Int,
        left:Int,
        bottom:Int,
        right:Int,
    )
    extends RectangleImpl(sheet, top, left, bottom, right)
    with Table[RectangleImpl]
    with StackedTableQuery[RectangleImpl]
    with RectangleLineDraw {

    val tableFunc = TableFunctionImpl
    val createTableQuery = (rect:RectangleImpl) => new TableQueryImpl(
        rect.sheet, rect.top, rect.left, rect.bottom, rect.right)
}

object TableQueryImpl {
    implicit def apply(
        sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int) =
            new TableQueryImpl(sheet,top,left,bottom,right)
    implicit def apply(rect:RectangleImpl) = 
            new TableQueryImpl(
                rect.sheet,rect.top,rect.left,rect.bottom,rect.right)
}


trait RectangleLineDraw {
    rect:RectangleImpl =>

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

    def getTableMap()(implicit newInstance:(RectangleImpl)=>TableQueryImpl)
            :Map[String,TableQueryImpl] = {
        val tableList = sheet.getRectangleList[TableQueryImpl]

        tableList.map(t=>t.tableFunc.getTableName(t)).zipWithIndex.map(
            _ match {
                case ((Some(name), t), _) => (name, newInstance(t))
                case ((None, t), idx) => ("Table" + idx, newInstance(t))
            }).toMap
    }
}
