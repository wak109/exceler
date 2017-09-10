/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xls

import scala.collection._
import scala.language.implicitConversions

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.common._
import exceler.xml._
import exceler.excel._
import exceler.rect.Table

import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._

import exceler.xml._

object XlsTable {

  def apply(sheet:Sheet, top:Int, left:Int, height:Int, width:Int) = {
    val topLeftList = getTopLeftList(sheet,top,left,height,width)
    val xlsRowLineList = getXlsRowLineList(topLeftList)
    val xlsColumnLineList = getXlsColumnLineList(topLeftList)
    val xlsRectList = topLeftList.map(t => XlsRect(
        sheet, t._1, t._2,
        getXlsHeight(sheet, t._1, t._2, top + height),
        getXlsWidth(sheet, t._1, t._2, left + width)))
    val xmlRectList = xlsRectList.map(t => getXmlRect( 
        t.top, t.left, t.height, t.width,
        xlsRowLineList, xlsColumnLineList))
    val xlsCellList = xlsRectList.zip(xmlRectList).map(
        t => new XlsCell(t._1, t._2))

    Table(xlsCellList)
  }

  def getTopLeftList(
      sheet:Sheet, top:Int, left:Int, height:Int, width:Int) =
    for {
      row <- (top until top + height)
      col <- (left until left + width)
      cell = sheet.cell(row, col)
        if (cell.hasBorderTop || row == top) &&
           (cell.hasBorderLeft || col == left)
    } yield (row, col)

  def getXlsHeight(sheet:Sheet, top:Int, col:Int, tableBottom:Int) =
    (for {
      row <- (top until tableBottom)
      cell = sheet.cell(row, col) if cell.hasBorderBottom
    } yield row - top + 1).headOption.getOrElse(tableBottom - top)

  def getXlsWidth(sheet:Sheet, row:Int, left:Int, tableRight:Int) =
    (for {
      col <- (left until tableRight)
      cell = sheet.cell(row, col) if cell.hasBorderRight
    } yield col - left + 1).headOption.getOrElse(tableRight - left)

  def getXlsRowLineList(topLeftList:Seq[(Int,Int)]) =
    topLeftList.map(_._1).toSet.toList.sorted

  def getXlsColumnLineList(topLeftList:Seq[(Int,Int)]) =
    topLeftList.map(_._2).toSet.toList.sorted

  def getXmlRowSpan(top:Int, height:Int, hLineList:Seq[Int]) = {
    val xmlRowList = hLineList.zipWithIndex.dropWhile(_._1 < top)
      .takeWhile(_._1 < top + height).map(_._2)

    (xmlRowList.head, xmlRowList.length)
  }

  def getXmlColumnSpan(left:Int, width:Int, vLineList:Seq[Int]) = {
    val xmlColumnList = vLineList.zipWithIndex.dropWhile(_._1 < left)
      .takeWhile(_._1 < left + width).map(_._2)

    (xmlColumnList.head, xmlColumnList.length)
  }

  def getXmlRect(top:Int, left:Int, height:Int, width:Int, 
      hLineList:Seq[Int], vLineList:Seq[Int]) = {
    val xmlRowSpan = getXmlRowSpan(top, height, hLineList)
    val xmlColumnSpan = getXmlColumnSpan(left, width, vLineList)

    (xmlRowSpan._1, xmlColumnSpan._1, xmlRowSpan._2, xmlColumnSpan._2)
  }
}
