/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

package ExcelLib.Rectangle {
    trait ImplicitConversions {
        implicit class ToExcelRectangleSheet(val sheet:Sheet)
            extends ExcelRectangleSheetConversion
    }
    object ImplicitConversions extends ImplicitConversions
}

import ExcelLib.ImplicitConversions._

trait ExcelRectangleSheetConversion {
    val sheet:Sheet

    import ExcelRectangleSheetConversion.Helper

    def getRectangleList():List[ExcelRectangle] = {
        for {
            cell <- Helper.getCellList(sheet)
            if cell.isOuterBorderBottom && cell.isOuterBorderRight
            topLeft <- Helper.findTopLeftFromBottomRight(cell)
        } yield new ExcelRectangle(sheet,
            topLeft.getRowIndex, topLeft.getColumnIndex,
            cell.getRowIndex, cell.getColumnIndex)
    }
}

object ExcelRectangleSheetConversion {

    object Helper {
    
        def findTopRightFromBottomRight(cell:Cell):Option[Cell] = (
            for {
                cOpt <- cell.getUpperStream
                c <- cOpt
                if c.isOuterBorderTop && c.isOuterBorderRight
            } yield c
        ).headOption
    
        def findBottomLeftFromBottomRight(cell:Cell):Option[Cell] = (
            for {
                cOpt <- cell.getLeftStream
                c <- cOpt
                if c.isOuterBorderBottom && c.isOuterBorderLeft
            } yield c
        ).headOption
    
        def findTopLeftFromBottomRight(cell:Cell):Option[Cell] = {
            (findTopRightFromBottomRight(cell), 
                    findBottomLeftFromBottomRight(cell)) match {
                case (Some(topRight), Some(bottomLeft)) =>
                    cell.getSheet.getCellOption(
                        topRight.getRowIndex, bottomLeft.getColumnIndex)
                case _ => None
            }
        }
    
        def getCellList(sheet:Sheet):List[Cell] = {
            for {
                row <- for {
                    rownum <- (sheet.getFirstRowNum
                            to sheet.getLastRowNum).toList
                    row <- sheet.getRowOption(rownum)
                } yield row
                colnum <- (row.getFirstCellNum
                            until row.getLastCellNum).toList
                cell <- row.getCellOption(colnum)
            } yield cell
        }
    }
}
