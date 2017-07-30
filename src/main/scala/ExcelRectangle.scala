/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook


trait ExcelRectangle {
    val sheet:Sheet
    val top:Int
    val left:Int
    val bottom:Int
    val right:Int
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
}
