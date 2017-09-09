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
import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._

package excellib.Table {
  trait ImplicitConversions {
    implicit class ToExcelTableSheetConversion(val sheet:Sheet)
        extends ExcelTableSheetConversion
  }
  object ImplicitConversions extends ImplicitConversions
}


trait ExcelFactory[T] {
  def create(sheet:Sheet,top:Int,left:Int,bottom:Int,right:Int):T
  def create(rect:ExcelRectangle):T = this.create(
    rect.sheet, rect.top, rect.left, rect.bottom, rect.right)
}

trait ExcelTableSheetFunction extends ExcelTableFunction {

  override val tableFunction = new SheetFunctionImpl{}

  trait SheetFunctionImpl extends FunctionImpl {

    def getTableName(rect:ExcelRectangle)
        : (Option[String], ExcelRectangle) = {
      // Table name outside of Rectangle
      (rect.top match {
        case 0 => (None, rect)
        case _ => (rect.sheet.cell(rect.top - 1, rect.left)
            .getValueString, rect)
      }) match {
        case (Some(name), _) => (Some(name), rect)
        case (None, _) => {
          // Table name at the top of Rectangle
          val (rowHead, rowTail) = tableFunction.getHeadRow(rect)
          val (rowHeadLeft, rowHeadRight) =
              tableFunction.getHeadCol(rowHead)
  
          rowHeadRight match {
            case Some(_) => (None, rect)
            case None => (tableFunction.getValue(rowHeadLeft),
                rowTail) match {
              case (name, Some(tail)) => (name, tail)
              case (name, None) => (None, rect)
            }
          }
        }
      }
    }
  }
}
  

trait ExcelTableSheetConversion extends ExcelTableSheetFunction {

  val sheet:Sheet

  def getTableMap[T:ExcelFactory]():Map[String,T] = {
    val factory = implicitly[ExcelFactory[T]]
    val tableList = sheet.getRectangleList

    tableList.map(t=>tableFunction.getTableName(t)).zipWithIndex.map(
      _ match {
        case ((Some(name), t), _) => (name, factory.create(t))
        case ((None, t), idx) => ("Table" + idx, factory.create(t))
      }).toMap
  }
}
