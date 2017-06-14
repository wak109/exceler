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

trait ExcelTableQuery[T <: ExcelTableQuery[T]] extends RectangleGrid[T] {
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
                row <- this.queryRow(pred)
                rtail = row.columnList.tail
                rowNext <- newInstance(
                    rtail.head.sheet,
                    rtail.head.topRow,
                    rtail.head.leftCol,
                    rtail.last.bottomRow,
                    rtail.last.rightCol
                ).queryRow(predTail)
            } yield rowNext
        }
    }

    def queryRow(pred:String => Boolean):List[T] = {
        for {
            row <- this.rowList
            if pred(row.columnList.head.value.getOrElse(""))
        } yield row
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
    extends RectangleGrid[ExcelTable]
    with ExcelTableQuery[ExcelTable]
    with ExcelNameAndTable {

    lazy val rowList = this.getRowList
    lazy val columnList = this.getColumnList
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

    override def toString():String =
        "ExcelTable:" + sheet.getSheetName + ":(" + 
            sheet.cell(topRow, leftCol).getAddress + "," +
            sheet.cell(bottomRow, rightCol).getAddress + ")"

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
