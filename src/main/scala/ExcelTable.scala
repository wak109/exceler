/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib._


class ExcelRectangle(
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
    ) {

    override def toString():String =
        "ExcelTable:" + sheet.getSheetName + ":(" + 
            sheet.cell_(topRow, leftCol).getAddress + "," +
            sheet.cell_(bottomRow, rightCol).getAddress + ")"
}

class ExcelTableCell(
    sheet:Sheet,
    topRow:Int,
    leftCol:Int,
    bottomRow:Int,
    rightCol:Int
    ) extends ExcelRectangle(sheet, topRow, leftCol, bottomRow, rightCol) {

}

class ExcelTable(
    sheet:Sheet,
    topRow:Int,
    leftCol:Int,
    bottomRow:Int,
    rightCol:Int
    ) extends ExcelRectangle(sheet, topRow, leftCol, bottomRow, rightCol) {


    def this(topLeft:Cell, bottomRight:Cell) = this(
        topLeft.getSheet,
        topLeft.getRowIndex,
        topLeft.getColumnIndex,
        bottomRight.getRowIndex,
        bottomRight.getColumnIndex
        )

    import ExcelTable._

    def getLeftColumnCellList:List[Int] = {
        for {
             rownum <- (topRow to bottomRow).toList
             cell <- sheet.getCell_(rownum, leftCol)
             if cell.hasBorderBottom_
        } yield cell.getRowIndex
    }

    println(getLeftColumnCellList)
}

object ExcelTable {


    implicit class ExcelTableCell(cell:Cell) {

        ////////////////////////////////////////////////////////////////
        // hasBorder
        //
        def hasBorderBottom_():Boolean = {
            (cell.getCellStyle.getBorderBottomEnum != BorderStyle.NONE) ||
                (cell.getLowerCell_.map(
                    _.getCellStyle.getBorderTopEnum != BorderStyle.NONE)
                match {
                    case Some(b) => b
                    case None => false
                })
        }

        def hasBorderTop_():Boolean = {
            (cell.getRowIndex == 0) ||
            (cell.getCellStyle.getBorderTopEnum != BorderStyle.NONE) ||
                (cell.getUpperCell_.map(
                    _.getCellStyle.getBorderBottomEnum != BorderStyle.NONE)
                match {
                    case Some(b) => b
                    case None => false
                })
        }

        def hasBorderRight_():Boolean = {
            (cell.getCellStyle.getBorderRightEnum != BorderStyle.NONE) ||
                (cell.getRightCell_.map(
                    _.getCellStyle.getBorderLeftEnum != BorderStyle.NONE)
                match {
                    case Some(b) => b
                    case None => false
                })
        }

        def hasBorderLeft_():Boolean = {
            (cell.getColumnIndex == 0) ||
            (cell.getCellStyle.getBorderLeftEnum != BorderStyle.NONE) ||
                (cell.getLeftCell_.map(
                    _.getCellStyle.getBorderRightEnum != BorderStyle.NONE)
                match {
                    case Some(b) => b
                    case None => false
                })
        }

        ////////////////////////////////////////////////////////////////
        // isOuterBorder
        //
        //
        def isOuterBorderTop_():Boolean = {
            val upperCell = cell.getUpperCell_

            (cell.hasBorderTop_) &&
            (upperCell match {
                case Some(cell) =>
                    (! cell.hasBorderLeft_) &&
                    (! cell.hasBorderRight_)
                case None => true
            })
        }

        def isOuterBorderBottom_():Boolean = {
            val lowerCell = cell.getLowerCell_

            (cell.hasBorderBottom_) &&
            (lowerCell match {
                case Some(cell) =>
                    (! cell.hasBorderLeft_) &&
                    (! cell.hasBorderRight_)
                case None => true
            })
        }

        def isOuterBorderLeft_():Boolean = {
            val leftCell = cell.getLeftCell_

            (cell.hasBorderLeft_) &&
            (leftCell match {
                case Some(cell) =>
                    (! cell.hasBorderTop_) &&
                    (! cell.hasBorderBottom_)
                case None => true
            })
        }

        def isOuterBorderRight_():Boolean = {
            val rightCell = cell.getRightCell_

            (cell.hasBorderRight_) &&
            (rightCell match {
                case Some(cell) =>
                    (! cell.hasBorderTop_) &&
                    (! cell.hasBorderBottom_)
                case None => true
            })
        }
    }

    ////////////////////////////////////////////////////////////////
    // Function
    //
    def findTopRightFromBottomRight(cell:Cell):Option[Cell] = (
        for {
            c <- cell.getUpperStream_
            if (c.map(_.isOuterBorderTop_) == Some(true)) &&
                (c.map(_.isOuterBorderRight_) == Some(true))
        } yield c.get
    ).headOption

    def findBottomLeftFromBottomRight(cell:Cell):Option[Cell] = (
        for {
            c <- cell.getLeftStream_
            if (c.map(_.isOuterBorderBottom_) == Some(true)) &&
                (c.map(_.isOuterBorderLeft_) == Some(true))
        } yield c.get
    ).headOption

    def findTopLeftFromBottomRight(cell:Cell):Option[Cell] = {
        (findTopRightFromBottomRight(cell), 
                findBottomLeftFromBottomRight(cell)) match {
            case (Some(topRight), Some(bottomLeft)) =>
                cell.getSheet.getCell_(
                    topRight.getRowIndex, bottomLeft.getColumnIndex)
            case _ => None
        }
    }

    def getCellList(sheet:Sheet):List[Cell] = {
        for {
             row <- (sheet.getFirstRowNum to
                 sheet.getLastRowNum).toList.map(sheet.getRow(_))
             if row != null
             cell <- (row.getFirstCellNum until
                 row.getLastCellNum).toList.map(row.getCell(_))
             if cell != null
        } yield cell
    }

    def getExcelTableList(sheet:Sheet):List[ExcelTable] = {
        for {
            cell <- getCellList(sheet)
            if cell.isOuterBorderBottom_ && cell.isOuterBorderRight_
            topLeft = findTopLeftFromBottomRight(cell)
            if topLeft.nonEmpty
        } yield new ExcelTable(topLeft.get, cell)
    }
}
