/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib._
import ExcelRectangle._

abstract trait ExcelTableQuery[T <: ExcelTableQuery[T]] {
    this:T =>

    val sheet:Sheet
    val topRow:Int
    val leftCol:Int
    val bottomRow:Int
    val rightCol:Int

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
                    rtail(rtail.length - 1).bottomRow,
                    rtail(rtail.length - 1).rightCol
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
                    ctail(ctail.length - 1).bottomRow,
                    ctail(ctail.length - 1).rightCol
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

class ExcelTable (
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
    )
    extends ExcelRectangle[ExcelTable]
    with ExcelTableQuery[ExcelTable] {

    lazy val rowList = this.getRowList
    lazy val columnList = this.getColumnList
    lazy val value = this.getSingleValue

    def this(t:ExcelTable) = this(
            t.sheet, t.topRow, t.leftCol, t.bottomRow, t.rightCol)

    def getSingleValue():Option[String] = (
        for {
            colnum <- (leftCol to rightCol).toStream
            rownum <- (topRow to bottomRow).toStream
            value <- sheet.cell(rownum, colnum).getValueString.map(_.trim)
        } yield value
    ).headOption

    def getTableName():Option[String] = {
        topRow match {
            case 0 => None
            case _ => sheet.cell(topRow - 1, leftCol).getValueString
        }
    }

    override def toString():String =
        "ExcelTable:" + sheet.getSheetName + ":(" + 
            sheet.cell(topRow, leftCol).getAddress + "," +
            sheet.cell(bottomRow, rightCol).getAddress + ")"

}

object ExcelTable {

    implicit def newExcelTable(
        sheet:Sheet,
        topRow:Int,
        leftCol:Int,
        bottomRow:Int,
        rightCol:Int
    ):ExcelTable =
        new ExcelTable(sheet, topRow, leftCol, bottomRow, rightCol)

    implicit class ExcelTableSheetExtra (sheet:Sheet) {

        def getTableMap():Map[String,ExcelTable] = {
            val tableList = sheet.getRectangleList[ExcelTable]
            tableList.zip(tableList.map(_.getTableName)).zipWithIndex.map(
                _ match {
                    case ((t, Some(name)), _) => (name, t)
                    case ((t, None), idx) => ("Table" + idx, t)
                }).toMap
        }
    }
}
