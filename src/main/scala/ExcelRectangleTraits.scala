/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._

import java.io._
import java.nio.file._

import ExcelLib._
import ExcelRectangle._

trait ExcelRectangleDraw {

    val sheet:Sheet
    val topRow:Int
    val leftCol:Int
    val bottomRow:Int
    val rightCol:Int

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
