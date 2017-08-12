/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
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


class ExcelTableQuery (val sheet:Sheet, val top:Int, val left:Int,
  val bottom:Int, val right:Int)
    extends TableQuery[ExcelRectangle]
    with ExcelRectangle
    with ExcelTableQueryFunction {
  override val tableQueryFunction = new QueryFunctionImpl{}
}


trait ExcelTableQueryFunction
  extends ExcelTableFunction 
  with TableQueryFunction[ExcelRectangle] {

  override val tableQueryFunction = new QueryFunctionImpl{}

  trait QueryFunctionImpl extends QueryFunction {

    override def create(rect:ExcelRectangle) =
      new ExcelTableQuery(rect.sheet, rect.top,
          rect.left, rect.bottom, rect.right)
  }
}
