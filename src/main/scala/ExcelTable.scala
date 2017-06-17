/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._

package ExcelLib.Table {
    trait ImplicitConversions {
        implicit class ToExcelTableSheetConversion(val sheet:Sheet)
                extends ExcelTableSheetConversion
    }
    object ImplicitConversions extends ImplicitConversions
}

trait Rectangle[T <: Rectangle[T]] {
    def getCross(col:T):T
}

trait TableCell {
    val value:String
}

/*
trait HeadTail[T] {
    def getRow(rect:T):(T,Option[T])
    def getCol(rect:T):(T,Option[T])
}
*/

trait Table[T] {
    rect:T =>
//    val headTail:HeadTail[T]

    lazy val rowList:List[T] = this.getRowList(Some(rect))
    lazy val colList:List[T] = this.getColList(Some(rect))

    def getRow(rect:T):(T,Option[T])
    def getCol(rect:T):(T,Option[T])

    def getRowList(rect:Option[T]):List[T] = {
        rect match {
            case None => Nil
            case Some(rect) => {
                val (headRow, tailRow) = this.getRow(rect)
                headRow :: getRowList(tailRow)
            }
        }
    }

    def getColList(rect:Option[T]):List[T] = {
        rect match {
            case None => Nil
            case Some(rect) => {
                val (headCol, tailCol) = this.getCol(rect)
                headCol :: getColList(tailCol)
            }
        }
    }
}

trait TableQuery[T <: Rectangle[T]] extends Table[T] {
    rect:T =>
    val newQ:(T) => TableQuery[T]
    val newCell:(T) => TableCell

    def query(
        rowpredList:List[(String) => Boolean],
        colpredList:List[(String) => Boolean]
    ):List[List[T]] = {
        for {
            row <- this.queryRow(rowpredList)
        } yield for {
            col <- this.queryColumn(colpredList)
        } yield row.getCross(col)
    }

    def query(
        rowpred:(String) => Boolean,
        colpred:(String) => Boolean
    ):List[List[T]] = {
        for {
            row <- this.queryRow(rowpred)
        } yield for {
            col <- this.queryColumn(colpred)
        } yield row.getCross(col)
    }

    def queryRow(
        predList:List[String => Boolean]
    ):List[T] = {
        predList match {
            case Nil => List[T](rect)
            case pred::predTail => for {
                row <- this.queryRow(pred)
                row <- newQ(row).queryRow(predTail)
            } yield row
        }
    }

    def queryRow(
        pred:String => Boolean
    ):List[T] = {
        for {
            row <- this.rowList
            (colHead, colTail) = this.getCol(row)
            if pred(newCell(colHead).value)
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
                col <- newQ(col).queryColumn(predTail)
            } yield col
        }
    }

    def queryColumn(
        pred:String => Boolean
    ) :List[T] = {
        for {
            col <- this.colList
            (rowHead, rowTail) = this.getRow(col)
            if pred(newCell(rowHead).value)
            row <- rowTail
        } yield row
    }
}

/*
trait StackedTableQuery[T <: Rectangle[T]] extends TableQuery[T] {
    rect:T =>

    val tableMap:Map[String, TableQuery[T]]

    override def queryRow(
        predList:List[String => Boolean]
    )(implicit newQ:(T) => TableQuery[T],
            newCell:(T) => TableCell):List[T] = {

        predList match {
            case Nil => List[T](this)
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
*/
////////////////////////////////////////////////////////////////////////
// Impl
//

class RectangleImpl (
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
) extends Rectangle[RectangleImpl] {

    override def getCross(col:RectangleImpl) = {
        new RectangleImpl(this.sheet, this.topRow, col.leftCol,
                this.bottomRow, col.rightCol)
    }
}

object RectangleImpl {
    implicit def apply(
        sheet:Sheet,
        topRow:Int,
        leftCol:Int,
        bottomRow:Int,
        rightCol:Int
        ) = new RectangleImpl(sheet, topRow, leftCol, bottomRow, rightCol)
}

class TableCellImpl(val rect:RectangleImpl) extends TableCell {

    lazy val value = this.getSingleValue().getOrElse("")

    def getSingleValue():Option[String] = (
        for {
            colnum <- (rect.leftCol to rect.rightCol).toStream
            rownum <- (rect.topRow to rect.bottomRow).toStream
            value <- rect.sheet.cell(rownum, colnum)
                        .getValueString.map(_.trim)
        } yield value
    ).headOption
}

object TableCellImpl {
    implicit def newTableCellImpl(rect:RectangleImpl) =
            new TableCellImpl(rect)
}

trait TableImpl extends Table[RectangleImpl] {
    self:RectangleImpl =>
    
    override def getRow(rect:RectangleImpl) = {
        (for {
            rownum <- (rect.topRow until rect.bottomRow).toStream
            cell <- rect.sheet.getCellOption(rownum, rect.leftCol)
            if cell.hasBorderBottom
        } yield rownum).headOption match {
            case Some(num) => (
                new RectangleImpl(rect.sheet, rect.topRow,
                    rect.leftCol, num, rect.rightCol),
                Some(new RectangleImpl(rect.sheet, num + 1,
                    rect.leftCol, rect.bottomRow, rect.rightCol)))
            case _ => (rect, None)
        }
    }

    override def getCol(rect:RectangleImpl) = {
        (for {
            colnum <- (rect.leftCol until rect.rightCol).toStream
            cell <- rect.sheet.getCellOption(rect.topRow, colnum)
            if cell.hasBorderRight
        } yield colnum).headOption match {
            case Some(num) => (
                new RectangleImpl(rect.sheet, rect.topRow,
                    rect.leftCol, rect.bottomRow, num),
                Some(new RectangleImpl(rect.sheet, rect.topRow,
                    num + 1, rect.bottomRow, rect.rightCol)))
            case _ => (rect, None)
        }
    }
}

class TableQueryImpl(
        rect:RectangleImpl
    )(implicit
        _newQ:(RectangleImpl) => TableQueryImpl,
        _newCell:(RectangleImpl) => TableCellImpl
    )
    extends RectangleImpl(
        rect.sheet,
        rect.topRow,
        rect.leftCol,
        rect.bottomRow,
        rect.rightCol
    )
    with TableQuery[RectangleImpl]
    with TableImpl
    with RectangleBorderDraw {

    override val newQ = _newQ
    override val newCell = _newCell
}

object TableQueryImpl {
    implicit def apply(t:RectangleImpl):TableQueryImpl =
        new TableQueryImpl(t)
}

////////////////////////////////////////////////////////////////////////
//
trait ExcelTableQuery[T <: ExcelTableQuery[T]] extends RectangleGrid {
    this:T =>

    val rowList:List[T]
    val columnList:List[T]
    val value:Option[String]

    def query(
        rowpredList:List[String => Boolean],
        colpredList:List[String => Boolean]
        )(implicit newInstance:(
            Sheet,Int,Int,Int,Int)=>T):List[List[T]] = {
        for {
            row <- this.queryRow(rowpredList)
        } yield for {
            col <- this.queryColumn(colpredList)
        } yield newInstance(sheet,
            row.topRow, col.leftCol, row.bottomRow, col.rightCol)
    }

    def queryRow(predList:List[String => Boolean])(
            implicit newInstance:(Sheet,Int,Int,Int,Int)=>T):List[T] = {
        predList match {
            case Nil => List(this)
            case pred::predTail => for {
                rtail <- this.queryRow(pred)
                rowNext <- rtail.queryRow(predTail)
            } yield rowNext
        }
    }

    def queryRow(pred:String => Boolean)(
            implicit newInstance:(Sheet,Int,Int,Int,Int)=>T):List[T] = {
        for {
            rownum <- (0 until this.rowList.length).toList
            row = this.rowList(rownum)
            if pred(row.columnList.head.value.getOrElse(""))
        } yield {
            val rtail = if (row.columnList.length < 2) 
                (rownum + 1 until this.rowList.length)
                    .toList.map(this.rowList.apply(_))
            else
                row.columnList.tail

            newInstance(
                    rtail.head.sheet,
                    rtail.head.topRow,
                    rtail.head.leftCol,
                    rtail.last.bottomRow,
                    rtail.last.rightCol
            )
        }
    }

    def queryColumn(predList:List[String => Boolean])(
            implicit newInstance:(Sheet,Int,Int,Int,Int)=>T):List[T] = {
        predList match {
            case Nil => List(this)
            case pred::predTail => for {
                col <- this.queryColumn(pred)
                ctail = col.rowList.tail
                colNext <- newInstance(
                    ctail.head.sheet,
                    ctail.head.topRow,
                    ctail.head.leftCol,
                    ctail.last.bottomRow,
                    ctail.last.rightCol
                ).queryColumn(predTail)
            } yield colNext
        }
    }

    def queryColumn(pred:String => Boolean):List[T] = {
        for {
            col <- this.columnList
            if pred(col.rowList.head.value.getOrElse(""))
        } yield col
    }
}

trait ExcelNameAndTable {
    table:ExcelTable =>

    def getNameAndTable():(Option[String], ExcelTable) = {
        (topRow match {
            case 0 => (None, table)
            case _ => (sheet.cell(topRow - 1, leftCol).getValueString,
                            table)
        }) match {
            case (Some(name), t) => (Some(name), t)
            case (None, t) => t.rowList(0).columnList.length match {
                case r if r <= 1 => (t.rowList(0).getSingleValue,
                            new ExcelTable(t.rowList.tail))
                case _ => (None, t)
            }
        }
    }
}

case class ExcelTable (
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
    )
    extends RectangleGrid
    with ExcelTableQuery[ExcelTable]
    with ExcelNameAndTable {

    lazy val rowList = this.getRowList[ExcelTable]
    lazy val columnList = this.getColumnList[ExcelTable]
    lazy val value = this.getSingleValue()

    def this(t:ExcelTable) = this(
            t.sheet, t.topRow, t.leftCol, t.bottomRow, t.rightCol)

    def this(tList:List[ExcelTable]) = this(
        tList.head.sheet,
        tList.head.topRow,
        tList.head.leftCol,
        tList.last.bottomRow,
        tList.last.rightCol)

    def getSingleValue():Option[String] = (
        for {
            colnum <- (leftCol to rightCol).toStream
            rownum <- (topRow to bottomRow).toStream
            value <- sheet.cell(rownum, colnum).getValueString.map(_.trim)
        } yield value
    ).headOption
}

object ExcelTable {

    implicit def applyImplicitly(
        sheet:Sheet,
        topRow:Int,
        leftCol:Int,
        bottomRow:Int,
        rightCol:Int
    ):ExcelTable = this.apply(sheet, topRow, leftCol, bottomRow, rightCol)
}


trait ExcelTableSheetConversion {
    val sheet:Sheet

    def getTableMap:Map[String,ExcelTable] = {
        val tableList = sheet.getRectangleList[ExcelTable]

        tableList.map(_.getNameAndTable).zipWithIndex.map(
            _ match {
                case ((Some(name), t), _) => (name, t)
                case ((None, t), idx) => ("Table" + idx, t)
            }).toMap
    }
}
