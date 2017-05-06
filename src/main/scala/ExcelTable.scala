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

    override def toString():String =
        "ExcelTableCell:" + sheet.getSheetName + ":(" + 
            sheet.cell_(topRow, leftCol).getAddress + "," +
            sheet.cell_(bottomRow, rightCol).getAddress + ")"
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

    import ExcelTable.CellImplicitForExcelTable

    val tableCell = {
        val rownumList = (topRow-1) :: (for {
             rownum <- (topRow to bottomRow).toList
             cell <- sheet.getCell_(rownum, leftCol)
             if cell.hasBorderBottom_
        } yield rownum)

        val colnumList = (leftCol-1) :: (for {
             colnum <- (leftCol to rightCol).toList
             cell <- sheet.getCell_(topRow, colnum)
             if cell.hasBorderRight_
        } yield colnum)

        val tCellArray = Array.ofDim[ExcelTableCell](
            rownumList.length - 1, colnumList.length -1)

        for {
            (trowidx, rownumStart, rownumEnd)
                <- (Stream.from(0), rownumList, rownumList.drop(1)).zipped
            (tcolidx, colnumStart, colnumEnd)
                <- (Stream.from(0), colnumList, colnumList.drop(1)).zipped
        } {
            tCellArray(trowidx)(tcolidx) =
                new ExcelTableCell(sheet,
                    rownumStart + 1, colnumStart + 1, rownumEnd, colnumEnd)

            println(tCellArray(trowidx)(tcolidx))
        }
    }
}

object ExcelTable {


    implicit class CellImplicitForExcelTable (cell:Cell) {

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
            cOpt <- cell.getUpperStream_
            c <- cOpt
            if c.isOuterBorderTop_ && c.isOuterBorderRight_
        } yield c
    ).headOption

    def findBottomLeftFromBottomRight(cell:Cell):Option[Cell] = (
        for {
            cOpt <- cell.getLeftStream_
            c <- cOpt
            if c.isOuterBorderBottom_ && c.isOuterBorderLeft_
        } yield c
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
            row <- for {
                rownum <- (sheet.getFirstRowNum to sheet.getLastRowNum).toList
                row <- sheet.getRow_(rownum)
            } yield row
            colnum <- (row.getFirstCellNum until row.getLastCellNum).toList
            cell <- row.getCell_(colnum)
        } yield cell
    }

    def getExcelTableList(sheet:Sheet):List[ExcelTable] = {
        for {
            cell <- getCellList(sheet)
            if cell.isOuterBorderBottom_ && cell.isOuterBorderRight_
            topLeft <- findTopLeftFromBottomRight(cell)
        } yield new ExcelTable(topLeft, cell)
    }
}
