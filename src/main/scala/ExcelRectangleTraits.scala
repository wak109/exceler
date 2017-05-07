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
            sheet.cell_(topRow, colnum).setBorderTop_(borderStyle)
    }

    def drawOuterBorderLeft(borderStyle:BorderStyle):Unit = {
        for (rownum <- (topRow to bottomRow).toList)
            sheet.cell_(rownum, leftCol).setBorderLeft_(borderStyle)
    }

    def drawOuterBorderBottom(borderStyle:BorderStyle):Unit = {
        for (colnum <- (leftCol to rightCol).toList)
            sheet.cell_(bottomRow, colnum).setBorderBottom_(borderStyle)
    }

    def drawOuterBorderRight(borderStyle:BorderStyle):Unit = {
        for (rownum <- (topRow to bottomRow).toList)
            sheet.cell_(rownum, rightCol).setBorderRight_(borderStyle)
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
                sheet.cell_(0, colnum).setBorderTop_(borderStyle)
        }
        else if (0 < num && num <= bottomRow - topRow) {
            for (colnum <- (leftCol to rightCol).toList)
                sheet.cell_(topRow + num - 1, colnum
                    ).setBorderBottom_(borderStyle)
        }
    }

    def drawVerticalLine(num:Int, borderStyle:BorderStyle):Unit = {
        if (num == 0) {
            for (rownum <- (topRow to bottomRow).toList)
                sheet.cell_(rownum, 0).setBorderLeft_(borderStyle)
        }
        else if (0 < num && num <= rightCol - leftCol) {
            for (rownum <- (topRow to bottomRow).toList)
                sheet.cell_(rownum, leftCol + num - 1
                    ).setBorderRight_(borderStyle)
        }
    }
}
