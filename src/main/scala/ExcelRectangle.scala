/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

package ExcelLib.Rectangle {
    trait ImplicitConversions {
        implicit class ToExcelRectangleSheet(val sheet:Sheet)
            extends ExcelRectangleSheetConversion
    }
    object ImplicitConversions extends ImplicitConversions
}

import ExcelLib.ImplicitConversions._

abstract class ExcelRectangle {
    val sheet:Sheet
    val topRow:Int
    val leftCol:Int
    val bottomRow:Int
    val rightCol:Int
}

trait RectangleGrid[T <: ExcelRectangle] extends ExcelRectangle {

    def getRowList()(implicit newInstance:(
            Sheet, Int, Int, Int, Int) => T):List[T] = {
        val rownumList = (topRow-1) :: (for {
            rownum <- (topRow until bottomRow).toList
            cell <- sheet.getCellOption(rownum, leftCol)
            if cell.hasBorderBottom
        } yield rownum) ::: List(bottomRow)

        (rownumList, rownumList.drop(1)).zipped.toList.map(
                tpl => newInstance(
                    sheet, tpl._1 + 1, leftCol, tpl._2,  rightCol))
    }

    def getColumnList()(implicit newInstance:(
            Sheet, Int, Int, Int, Int) => T):List[T] = {
        val colnumList = (leftCol-1) :: (for {
            colnum <- (leftCol until rightCol).toList
            cell <- sheet.getCellOption(topRow, colnum)
            if cell.hasBorderRight
        } yield colnum) ::: List(rightCol)

        (colnumList, colnumList.drop(1)).zipped.toList.map(
                tpl => newInstance(
                    sheet, topRow, tpl._1 + 1, bottomRow, tpl._2))
    }

    override def toString():String =
        this.getClass.getSimpleName + ":" + sheet.getSheetName + ":(" +
            sheet.cell(topRow, leftCol).getAddress + "," +
            sheet.cell(bottomRow, rightCol).getAddress + ")"
}



trait RectangleBorderDraw extends ExcelRectangle {

    def drawOuterBorderTop(borderStyle:BorderStyle):Unit = {
        for (colnum <- (leftCol to rightCol).toList)
            sheet.cell(topRow, colnum).setBorderTop(borderStyle)
    }

    def drawOuterBorderLeft(borderStyle:BorderStyle):Unit = {
        for (rownum <- (topRow to bottomRow).toList)
            sheet.cell(rownum, leftCol).setBorderLeft(borderStyle)
    }

    def drawOuterBorderBottom(borderStyle:BorderStyle):Unit = {
        for (colnum <- (leftCol to rightCol).toList)
            sheet.cell(bottomRow, colnum).setBorderBottom(borderStyle)
    }

    def drawOuterBorderRight(borderStyle:BorderStyle):Unit = {
        for (rownum <- (topRow to bottomRow).toList)
            sheet.cell(rownum, rightCol).setBorderRight(borderStyle)
    }

    def drawOuterBorder(borderStyle:BorderStyle):Unit = {
        drawOuterBorderTop(borderStyle)
        drawOuterBorderLeft(borderStyle)
        drawOuterBorderBottom(borderStyle)
        drawOuterBorderRight(borderStyle)
    }

    def drawHorizontalLine(num:Int, borderStyle:BorderStyle):Unit = {
        if (num == 0) {
            for (colnum <- (leftCol to rightCol).toList)
                sheet.cell(0, colnum).setBorderTop(borderStyle)
        }
        else if (0 < num && num <= bottomRow - topRow) {
            for (colnum <- (leftCol to rightCol).toList)
                sheet.cell(topRow + num - 1, colnum
                    ).setBorderBottom(borderStyle)
        }
    }

    def drawVerticalLine(num:Int, borderStyle:BorderStyle):Unit = {
        if (num == 0) {
            for (rownum <- (topRow to bottomRow).toList)
                sheet.cell(rownum, 0).setBorderLeft(borderStyle)
        }
        else if (0 < num && num <= rightCol - leftCol) {
            for (rownum <- (topRow to bottomRow).toList)
                sheet.cell(rownum, leftCol + num - 1
                    ).setBorderRight(borderStyle)
        }
    }
}


trait ExcelRectangleSheetConversion {
    val sheet:Sheet

    import ExcelRectangleSheetConversion.Helper

    def getRectangleList[T <: RectangleGrid[T]]()
            (implicit newInstance:(
                Sheet, Int, Int, Int, Int) => T):List[T] = {
        for {
            cell <- Helper.getCellList(sheet)
            if cell.isOuterBorderBottom && cell.isOuterBorderRight
            topLeft <- Helper.findTopLeftFromBottomRight(cell)
        } yield newInstance(sheet,
            topLeft.getRowIndex, topLeft.getColumnIndex,
            cell.getRowIndex, cell.getColumnIndex)
    }
}

object ExcelRectangleSheetConversion {

    object Helper {
    
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
                    rownum <- (sheet.getFirstRowNum
                            to sheet.getLastRowNum).toList
                    row <- sheet.getRowOption(rownum)
                } yield row
                colnum <- (row.getFirstCellNum
                            until row.getLastCellNum).toList
                cell <- row.getCellOption(colnum)
            } yield cell
        }
    }
}

