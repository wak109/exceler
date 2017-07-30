/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import CommonLib.ImplicitConversions._
import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._

package ExcelLib.Table {
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


trait ExcelTableSheetConversion
    extends ExcelTableFunction
    with ExcelTableQueryFunction {

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
