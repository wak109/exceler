/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelLib._
import ExcelRectangle._


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
        rect.rightCol
    )

    def getInnerRectangleList():List[List[ExcelRectangle]] = {

        val rownumList = (topRow-1) :: (for {
             rownum <- (topRow until bottomRow).toList
             cell <- sheet.getCell_(rownum, leftCol)
             if cell.hasBorderBottom
        } yield rownum) ::: List(bottomRow)

        val colnumList = (leftCol-1) :: (for {
             colnum <- (leftCol until rightCol).toList
             cell <- sheet.getCell_(topRow, colnum)
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
            sheet.cell_(topRow, leftCol).getAddress + "," +
            sheet.cell_(bottomRow, rightCol).getAddress + ")"
}

object ExcelRectangle {

        implicit class CellBorderImplicit (cell:Cell) {
    
            ////////////////////////////////////////////////////////////
            // hasBorder
            //
            def hasBorderBottom():Boolean = {
                (cell.getCellStyle.getBorderBottomEnum != 
                    BorderStyle.NONE) ||
                    (cell.getLowerCell.map(
                        _.getCellStyle.getBorderTopEnum != BorderStyle.NONE)
                    match {
                        case Some(b) => b
                        case None => false
                    })
            }
    
            def hasBorderTop():Boolean = {
                (cell.getRowIndex == 0) ||
                (cell.getCellStyle.getBorderTopEnum != BorderStyle.NONE) ||
                    (cell.getUpperCell.map(
                        _.getCellStyle.getBorderBottomEnum !=
                            BorderStyle.NONE)
                    match {
                        case Some(b) => b
                        case None => false
                    })
            }
    
            def hasBorderRight():Boolean = {
                (cell.getCellStyle.getBorderRightEnum !=
                    BorderStyle.NONE) ||
                    (cell.getRightCell.map(
                        _.getCellStyle.getBorderLeftEnum !=
                            BorderStyle.NONE)
                    match {
                        case Some(b) => b
                        case None => false
                    })
            }
    
            def hasBorderLeft():Boolean = {
                (cell.getColumnIndex == 0) ||
                (cell.getCellStyle.getBorderLeftEnum !=
                    BorderStyle.NONE) ||
                    (cell.getLeftCell.map(
                        _.getCellStyle.getBorderRightEnum !=
                            BorderStyle.NONE)
                    match {
                        case Some(b) => b
                        case None => false
                    })
            }
    
            ////////////////////////////////////////////////////////////
            // isOuterBorder
            //
            //
            def isOuterBorderTop():Boolean = {
                val upperCell = cell.getUpperCell
    
                (cell.hasBorderTop) &&
                (upperCell match {
                    case Some(cell) =>
                        (! cell.hasBorderLeft) &&
                        (! cell.hasBorderRight)
                    case None => true
                })
            }
    
            def isOuterBorderBottom():Boolean = {
                val lowerCell = cell.getLowerCell
    
                (cell.hasBorderBottom) &&
                (lowerCell match {
                    case Some(cell) =>
                        (! cell.hasBorderLeft) &&
                        (! cell.hasBorderRight)
                    case None => true
                })
            }
    
            def isOuterBorderLeft():Boolean = {
                val leftCell = cell.getLeftCell
    
                (cell.hasBorderLeft) &&
                (leftCell match {
                    case Some(cell) =>
                        (! cell.hasBorderTop) &&
                        (! cell.hasBorderBottom)
                    case None => true
                })
            }
    
            def isOuterBorderRight():Boolean = {
                val rightCell = cell.getRightCell
    
                (cell.hasBorderRight) &&
                (rightCell match {
                    case Some(cell) =>
                        (! cell.hasBorderTop) &&
                        (! cell.hasBorderBottom)
                    case None => true
                })
            }
        }

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

    }
}
