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

trait TableCell {
    val value:String
}

trait TableCellList[T <: TableCell] {
    def getHead():T
    def getTail():TableQuery[T]
}

trait Table[T <: TableCell] {
    val rowList:List[TableCellList[T]]
    val columnList:List[TableCellList[T]]
}

trait TableQuery[T <: TableCell] extends Table[T] {

    def query(
        rowpredList:List[String => Boolean],
        colpredList:List[String => Boolean]
        )(implicit newInstance:(TableQuery[T], TableQuery[T]) => T) = {
        for {
            row <- this.queryRow(rowpredList)
        } yield for {
            col <- this.queryColumn(colpredList)
        } yield newInstance(row, col)
    }

    def queryRow(predList:List[String => Boolean]):List[TableQuery[T]] = {
        predList match {
            case Nil => List[TableQuery[T]](this)
            case pred::predTail => for {
                row <- this.queryRow(pred)
                row <- row.queryRow(predTail)
            } yield row
        }
    }

    def queryRow(pred:String => Boolean):List[TableQuery[T]] = {
        for {
            row <- this.rowList
            if pred(row.getHead.value)
        } yield row.getTail
    }

    def queryColumn(predList:List[String => Boolean])
            :List[TableQuery[T]] = {
        predList match {
            case Nil => List[TableQuery[T]](this)
            case pred::predTail => for {
                col <- this.queryColumn(pred)
                col <- col.queryColumn(predTail)
            } yield col
        }
    }

    def queryColumn(pred:String => Boolean):List[TableQuery[T]] = {
        for {
            col <- this.columnList
            if pred(col.getHead.value)
        } yield col.getTail
    }
}

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
