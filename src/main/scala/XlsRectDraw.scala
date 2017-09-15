/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.excel._
import excellib.ImplicitConversions._


trait XlsTableDraw {

  /**
   * idx: 0 to height + 1
   */
  def drawHorizontalLine(rect:XlsRect,
      idx:Int, borderStyle:BorderStyle):Unit = {
    if (idx == 0) {
      for (col <- (rect.left to rect.right).toList)
        rect.sheet.cell(0, col).setBorderTop(borderStyle)
    }
    else if (0 < idx && idx <= rect.height) {
      for (col <- (rect.left to rect.right).toList)
        rect.sheet.cell(rect.top + idx - 1, col)
          .setBorderBottom(borderStyle)
    }
  }

  /**
   * idx: 0 to width + 1
   */
  def drawVerticalLine(rect:XlsRect,
      idx:Int, borderStyle:BorderStyle):Unit = {
    if (idx == 0) {
      for (row <- (rect.top to rect.bottom).toList)
        rect.sheet.cell(row, 0).setBorderLeft(borderStyle)
    }
    else if (0 < idx && idx <= rect.width) {
      for (row <- (rect.top to rect.bottom).toList)
        rect.sheet.cell(row, rect.left + idx - 1)
          .setBorderRight(borderStyle)
    }
  }

  def drawOuterBorderTop(rect:XlsRect, borderStyle:BorderStyle):Unit = {
    drawHorizontalLine(rect, 0, borderStyle)
  }

  def drawOuterBorderLeft(rect:XlsRect, borderStyle:BorderStyle):Unit = {
    drawVerticalLine(rect, 0, borderStyle)
  }

  def drawOuterBorderBottom(rect:XlsRect, borderStyle:BorderStyle):Unit = {
    drawHorizontalLine(rect, rect.height, borderStyle)
  }

  def drawOuterBorderRight(rect:XlsRect, borderStyle:BorderStyle):Unit = {
    drawVerticalLine(rect, rect.width, borderStyle)
  }

  def drawOuterBorder(rect:XlsRect, borderStyle:BorderStyle):Unit = {
    drawOuterBorderTop(rect, borderStyle)
    drawOuterBorderLeft(rect, borderStyle)
    drawOuterBorderBottom(rect, borderStyle)
    drawOuterBorderRight(rect, borderStyle)
  }
}
