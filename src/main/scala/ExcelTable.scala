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

//
// NormalTable
//
// table_name
// +-------+------+------+
// |       |      |      |    
// +-------+------+------+
//
//
// HeadingTable
//
// +---------------------+
// |  table_name         |
// +-------+------+------+
// |       |      |      |    
// +-------+------+------+

class ExcelTable (
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
    ) extends ExcelRectangle[ExcelTable] {

    def this(rect:ExcelTable) = this(
            rect.sheet,
            rect.topRow,
            rect.leftCol,
            rect.bottomRow,
            rect.rightCol)

    def this(rectList:List[ExcelTable]) = this(
            rectList.head.sheet,
            rectList.head.topRow,
            rectList.head.leftCol,
            rectList(rectList.length - 1).bottomRow,
            rectList(rectList.length - 1).rightCol)

    lazy val rowList = this.getRowList
    lazy val columnList = this.getColumnList
    lazy val value = this.getSingleValue

    def getSingleValue():Option[String] = (
        for {
            colnum <- (leftCol to rightCol).toStream
            rownum <- (topRow to bottomRow).toStream
            value <- sheet.cell(rownum, colnum).getValueString.map(_.trim)
        } yield value
    ).headOption

    def find(
        rowpredList:List[String => Boolean],
        colpredList:List[String => Boolean]
    ):List[List[ExcelTable]] = {
        for {
            row <- this.findRow(rowpredList)
        } yield {
            for {
                col <- this.findColumn(colpredList)
            } yield {
                new ExcelTable(sheet,
                    row.topRow, col.leftCol, row.bottomRow, col.rightCol)
            }
        }
    }

    def findRow(predList:List[String => Boolean]):List[ExcelTable] = {
        predList match {
            case Nil => List(this)
            case pred::predTail => for {
                row <- this.findRow(pred)
                rowNext <- (new ExcelTable(row.columnList.tail))
                        .findRow(predTail)
            } yield rowNext
        }
    }

    def findRow(pred:String => Boolean):List[ExcelTable] = {
        for {
            row <- this.rowList
            if pred(row.columnList.head.value.getOrElse(""))
        } yield row
    }

    def findColumn(predList:List[String => Boolean]):List[ExcelTable] = {
        predList match {
            case Nil => List(this)
            case pred::predTail => for {
                col <- this.findColumn(pred)
                colNext <- (new ExcelTable(col.rowList.tail))
                        .findColumn(predTail)
            } yield colNext
        }
    }

    def findColumn(pred:String => Boolean):List[ExcelTable] = {
        for {
            col <- this.columnList
            if pred(col.rowList.head.value.getOrElse(""))
        } yield col
    }

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
