/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._

trait RectangleLineDraw {
  rect:ExcelRectangle =>

  def drawOuterBorderTop(borderStyle:BorderStyle):Unit = {
    for (colnum <- (rect.left to rect.right).toList)
      rect.sheet.cell(rect.top, colnum).setBorderTop(borderStyle)
  }

  def drawOuterBorderLeft(borderStyle:BorderStyle):Unit = {
    for (rownum <- (rect.top to rect.bottom).toList)
      rect.sheet.cell(rownum, rect.left).setBorderLeft(borderStyle)
  }

  def drawOuterBorderBottom(borderStyle:BorderStyle):Unit = {
    for (colnum <- (rect.left to rect.right).toList)
      rect.sheet.cell(rect.bottom, colnum)
        .setBorderBottom(borderStyle)
  }

  def drawOuterBorderRight(borderStyle:BorderStyle):Unit = {
    for (rownum <- (rect.top to rect.bottom).toList)
      rect.sheet.cell(rownum, rect.right)
        .setBorderRight(borderStyle)
  }

  def drawOuterBorder(borderStyle:BorderStyle):Unit = {
    drawOuterBorderTop(borderStyle)
    drawOuterBorderLeft(borderStyle)
    drawOuterBorderBottom(borderStyle)
    drawOuterBorderRight(borderStyle)
  }

  def drawHorizontalLine(num:Int, borderStyle:BorderStyle):Unit = {
    if (num == 0) {
      for (colnum <- (rect.left to rect.right).toList)
        rect.sheet.cell(0, colnum).setBorderTop(borderStyle)
    }
    else if (0 < num && num <= rect.bottom - rect.top) {
      for (colnum <- (rect.left to rect.right).toList)
        rect.sheet.cell(rect.top + num - 1, colnum)
          .setBorderBottom(borderStyle)
    }
  }

  def drawVerticalLine(num:Int, borderStyle:BorderStyle):Unit = {
    if (num == 0) {
      for (rownum <- (rect.top to rect.bottom).toList)
        rect.sheet.cell(rownum, 0).setBorderLeft(borderStyle)
    }
    else if (0 < num && num <= rect.right - rect.left) {
      for (rownum <- (rect.top to rect.bottom).toList)
        rect.sheet.cell(rownum, rect.left + num - 1)
          .setBorderRight(borderStyle)
    }
  }
}
