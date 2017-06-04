/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._


import ExcelLib._


class ExcelRectangle (
    val sheet:Sheet,
    val topRow:Int,
    val leftCol:Int,
    val bottomRow:Int,
    val rightCol:Int
    ) {

    def this(rect:ExcelRectangle) = this(
            rect.sheet,
            rect.topRow,
            rect.leftCol,
            rect.bottomRow,
            rect.rightCol)

    def getInnerRectangleList():List[List[ExcelRectangle]] = {

        val rownumList = (topRow-1) :: (for {
             rownum <- (topRow until bottomRow).toList
             cell <- sheet.getCellOption(rownum, leftCol)
             if cell.hasBorderBottom
        } yield rownum) ::: List(bottomRow)

        val colnumList = (leftCol-1) :: (for {
             colnum <- (leftCol until rightCol).toList
             cell <- sheet.getCellOption(topRow, colnum)
             if cell.hasBorderRight
        } yield colnum) ::: List(rightCol)

        (for {
            (rownumStart, rownumEnd) <- (rownumList, rownumList.drop(1)).zipped
        } yield (for {
            (colnumStart, colnumEnd)
                <- (colnumList, colnumList.drop(1)).zipped
            } yield new ExcelRectangle(sheet,
                rownumStart + 1, colnumStart + 1, rownumEnd, colnumEnd)
            ).toList
        ).toList
    }

    override def toString():String =
        "ExcelRectangle:" + sheet.getSheetName + ":(" + 
            sheet.cell(topRow, leftCol).getAddress + "," +
            sheet.cell(bottomRow, rightCol).getAddress + ")"
}


object ExcelRectangle {

    implicit def excelRectangleImplicit(rect:ExcelRectangle) = (
            rect.sheet,
            rect.topRow,
            rect.leftCol,
            rect.bottomRow,
            rect.rightCol)

    implicit class SheetRectangleImplicit (sheet:Sheet) {

        def getRectangleList():List[ExcelRectangle] = {
            for {
                cell <- Helper.getCellList(sheet)
                if cell.isOuterBorderBottom && cell.isOuterBorderRight
                topLeft <- Helper.findTopLeftFromBottomRight(cell)
            } yield new ExcelRectangle(sheet,
                topLeft.getRowIndex, topLeft.getColumnIndex,
                cell.getRowIndex, cell.getColumnIndex)
        }
    }

    object Helper {
        ////////////////////////////////////////////////////////////////
        // Function
        //
        def findTopRightFromBottomRight(cell:Cell):Option[Cell] = (
            for {
                cOpt <- cell.getUpperStream
                c <- cOpt
                if c.isOuterBorderTop && c.isOuterBorderRight
            } yield c
        ).headOption
    
        def findBottomLeftFromBottomRight(cell:Cell):Option[Cell] = (
            for {
                cOpt <- cell.getLeftStream
                c <- cOpt
                if c.isOuterBorderBottom && c.isOuterBorderLeft
            } yield c
        ).headOption
    
        def findTopLeftFromBottomRight(cell:Cell):Option[Cell] = {
            (findTopRightFromBottomRight(cell), 
                    findBottomLeftFromBottomRight(cell)) match {
                case (Some(topRight), Some(bottomLeft)) =>
                    cell.getSheet.getCellOption(
                        topRight.getRowIndex, bottomLeft.getColumnIndex)
                case _ => None
            }
        }
    
        def getCellList(sheet:Sheet):List[Cell] = {
            for {
                row <- for {
                    rownum <- (sheet.getFirstRowNum to sheet.getLastRowNum).toList
                    row <- sheet.getRowOption(rownum)
                } yield row
                colnum <- (row.getFirstCellNum until row.getLastCellNum).toList
                cell <- row.getCellOption(colnum)
            } yield cell
        }

    }
}
