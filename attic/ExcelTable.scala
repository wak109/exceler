/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.excel

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import exceler.common._
import exceler.table._
import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._


trait ExcelTableFunction extends TableFunction[ExcelRectangle] {
  val tableFunction = new FunctionImpl

  class FunctionImpl extends Function {
    override def getCross(row:ExcelRectangle, col:ExcelRectangle) = 
      ExcelRectangle(row.sheet, row.top, col.left, row.bottom, col.right)
  
    override def getHeadRow(rect:ExcelRectangle):(
      ExcelRectangle, Option[ExcelRectangle]) = {
      (for {
        rownum <- (rect.top until rect.bottom).toStream
        cell <- rect.sheet.getCellOption(rownum, rect.left)
        if cell.hasBorderBottom
      } yield rownum).headOption match {
        case Some(num) => (
          ExcelRectangle(rect.sheet, rect.top,
            rect.left, num, rect.right),
          Some(ExcelRectangle(rect.sheet, num + 1,
            rect.left, rect.bottom, rect.right)))
        case _ => (rect, None)
      }
    }
  
    override def getHeadCol(rect:ExcelRectangle):(
        ExcelRectangle,Option[ExcelRectangle]) = {
      (for {
        colnum <- (rect.left until rect.right).toStream
        cell <- rect.sheet.getCellOption(rect.top, colnum)
        if cell.hasBorderRight
      } yield colnum).headOption match {
        case Some(num) => (
          ExcelRectangle(rect.sheet, rect.top,
            rect.left, rect.bottom, num),
          Some(ExcelRectangle(rect.sheet, rect.top,
            num + 1, rect.bottom, rect.right)))
        case _ => (rect, None)
      }
    }
  
    override def getValue(rect:ExcelRectangle):Option[String] =
      (for {
        colnum <- (rect.left to rect.right).toStream
        rownum <- (rect.top to rect.bottom).toStream
        value <- rect.sheet.cell(rownum, colnum)
              .getValueString.map(_.trim)
      } yield value).headOption
  
    override def mergeRect(rectL:List[ExcelRectangle]):ExcelRectangle = {
      val head = rectL.head
      val last = rectL.last
      ExcelRectangle(
        head.sheet, head.top, head.left, last.bottom, last.right)
    }
  }
}
