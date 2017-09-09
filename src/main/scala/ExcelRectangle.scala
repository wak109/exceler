/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.excel

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import excellib.ImplicitConversions._

trait ExcelRectangle {
  val sheet:Sheet
  val top:Int
  val left:Int
  val bottom:Int
  val right:Int

  def row(row:Int) = {
    assert((top <= row) && (row <= bottom))
    for (column <- Range(left, right + 1)) yield sheet.cell(row, column)
  }

  def column(column:Int) = {
    assert((left <= column) && (column <= right))
    for (row <- Range(top, bottom + 1)) yield sheet.cell(row, column)
  }
}


object ExcelRectangle {

  def apply(sheet:Sheet, top:Int, left:Int, bottom:Int, right:Int) = {

    class Impl (
      val sheet:Sheet,
      val top:Int,
      val left:Int,
      val bottom:Int,
      val right:Int
    ) extends ExcelRectangle

    new Impl(sheet, top, left, bottom, right)
  }

  implicit object factory extends ExcelFactory[ExcelRectangle] {
    override def create(
      sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int) =
        apply(sheet,top,left,bottom,right)
  }

  implicit object tableLib extends ExcelTableLibImpl {}
}

