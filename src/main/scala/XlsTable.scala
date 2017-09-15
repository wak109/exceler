/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.common._
import exceler.excel._

import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._

object XlsTable {

  def apply[T](sheet:Sheet, top:Int, left:Int, height:Int, width:Int)
  (implicit conv:T=>String,conv2:XlsRect=>T):Seq[Seq[XlsCell[T]]] = {
  
    val topLeftList = getTopLeftList(sheet,top,left,height,width)
    val xlsRowLineList = getXlsRowLineList(topLeftList)
    val xlsColumnLineList = getXlsColumnLineList(topLeftList)
    val xlsRectList = topLeftList.map(t => XlsRect(
        sheet, t._1, t._2,
        getXlsHeight(sheet, t._1, t._2, top + height),
        getXlsWidth(sheet, t._1, t._2, left + width)))
    val xmlRectList = xlsRectList.map(t => getXmlRect( 
        t.row, t.col, t.height, t.width,
        xlsRowLineList, xlsColumnLineList))
    val xlsCellList = xlsRectList.zip(xmlRectList).map(
        t => new XlsCell[T](t._1, t._2._1, t._2._2, t._2._3, t._2._4))

    TableX(xlsCellList)
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
  
  /**
   */
  def getRectList(sheet:Sheet, row:Int, col:Int,
      height:Int, width:Int):List[(Int,Int,Int,Int)] = {
    for {
      bottom <- (row until row + height).toList
      right <- (col until col + width).toList
      cell = sheet.cell(bottom, right)
        if cell.isOuterBorderBottom && cell.isOuterBorderRight
    } yield {
      val top = (for {
        row0 <- (bottom to row by -1).toList
        cell = sheet.cell(row0, right)
          if cell.isOuterBorderTop && cell.isOuterBorderRight
      } yield row0).headOption.getOrElse(row)
      val left = (for {
        col0 <- (right to col by -1).toList
        cell = sheet.cell(bottom, col0)
          if cell.isOuterBorderBottom && cell.isOuterBorderLeft
      } yield col0).headOption.getOrElse(col)

      (top, left, bottom - row + 1, right - col + 1)
    }
  }

  def getTableName(rect:XlsRect):Option[String] = {
    if (rect.row > 0) {
      for {
        cell <- rect.sheet.getCellOption(rect.row - 1, rect.col)
        value <- cell.getValueString
      } yield value
    } else None
  }

  def getTableNamePair[T](table:Seq[Seq[XlsCell[T]]])(
      implicit conv:T=>String):(Option[String],Seq[Seq[XlsCell[T]]]) = {
    val rect = table(0)(0).xlsRect

    getTableName(rect) match {
      case Some(name) => (Some(name), table)
      case None => table(0).length match {
        case 1 => (Some(table(0)(0).value), table.tail)
        case _ => (None, table)
      }
    }
  }

  def apply[T](sheet:Sheet)
  (implicit conv:T=>String, conv2:XlsRect=>T):Map[String,Seq[Seq[XlsCell[T]]]] = {
    (sheet.getUsedRange match {
      case Some((top, left, bottom, right)) => {
        val tableNamePairList:
          List[(Option[String],Seq[Seq[XlsCell[T]]])] = getRectList(
            sheet, top, left, bottom - top + 1, right - left +1)
              .map(t=>apply[T](sheet, t._1, t._2, t._3, t._4))
              .map(getTableNamePair[T](_))
        tableNamePairList.zipWithIndex.map(t=> {
          t._1._1 match {
            case Some(name) => (name, t._1._2)
            case None => ("Table" + t._2, t._1._2)
          }
        })
      }
      case None => Nil
    }).toMap
  }
}
