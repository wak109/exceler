/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

object ExcelLib {

        //
        // BorderTop
        // BorderBottom
        // BorderLeft
        // BorderRight
        // Foreground Color
        // Background Color
        // FillPatten
        // Horizontal Alignment
        // Vertical Alignment
        // Wrap Text
        //
    type CellStyleTuple = (
            BorderStyle, 
            BorderStyle,
            BorderStyle,
            BorderStyle,
            Color,
            Color,
            FillPatternType,
            HorizontalAlignment,
            VerticalAlignment,
            Boolean)

    implicit class WorkbookImplicit(workbook:Workbook)  {

        def saveAs_(filename:String): Unit =  {
            FileLib.createParentDir(filename)
            val out = new FileOutputStream(filename)
            workbook.write(out)
            out.close()
        }

        def getSheetOption(name:String):Option[Sheet] = Option(workbook.getSheet(name))

        def createSheet_(name:String):Try[Sheet] = Try(workbook.createSheet(name)) 

        def sheet(name:String):Sheet = {
            this.getSheetOption(name) match {
                case Some(s) => s
                case None => this.createSheet_(name) match {
                    case Success(s) => s
                    case Failure(e) => throw e
                }
            }
        }

        def findCellStyle(tuple:CellStyleTuple):Option[CellStyle] = (
            for {
                i <- (0 until workbook.getNumCellStyles).toStream
                cellStyle = workbook.getCellStyleAt(i)
                if cellStyle.toTuple == tuple
            } yield cellStyle
        ).headOption
    }

    implicit class SheetImplicit(sheet:Sheet) {

        def getRowOption(rownum:Int):Option[Row] = Option(sheet.getRow(rownum))

        def row(rownum:Int):Row = {
            this.getRowOption(rownum) match {
                case Some(r) => r
                case None => sheet.createRow(rownum)
            }
        }

        def getCellOption(rownum:Int, colnum:Int):Option[Cell] = 
            sheet.getRowOption(rownum).flatMap(_.getCellOption(colnum))

        def cell(rownum:Int, colnum:Int):Cell = sheet.row(rownum).cell(colnum)

        def getDrawingPatriarchOption():Option[Drawing[_ <: Shape]] = Option(sheet.getDrawingPatriarch)

        def drawingPatriarch():Drawing[_ <: Shape] = {
            this.getDrawingPatriarchOption match {
                case Some(d) => d
                case None => sheet.createDrawingPatriarch()
            }
        }
    }

    implicit class RowImplicit(row:Row) {

        def getCellOption(colnum:Int):Option[Cell] = Option(row.getCell(colnum))

        def cell(colnum:Int):Cell = row.getCell(colnum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
    }

    implicit class CellImplicit(cell:Cell) {

        def getValue_():Any = {
            cell.getCellTypeEnum match {
                case CellType.BLANK => cell.getStringCellValue
                case CellType.BOOLEAN => cell.getBooleanCellValue
                case CellType.ERROR => cell.getErrorCellValue
                case CellType.FORMULA => cell.getCellFormula
                case CellType.NUMERIC => cell.getNumericCellValue
                case CellType.STRING => cell.getStringCellValue
                case _ => cell.getStringCellValue
            }
        }

        def getUpperCell():Option[Cell] = {
            val rownum = cell.getRowIndex
            if (rownum > 0)
                cell.getSheet.getCellOption(rownum - 1, cell.getColumnIndex)
            else
                None
        }

        def getLowerCell():Option[Cell] =
            cell.getSheet.getCellOption(cell.getRowIndex + 1, cell.getColumnIndex)

        def getLeftCell():Option[Cell] = {
            val colnum = cell.getColumnIndex
            if (colnum > 0)
                cell.getSheet.getCellOption(cell.getRowIndex, colnum - 1)
            else
                None
        }

        def getRightCell():Option[Cell] =
            cell.getSheet.getCellOption(cell.getRowIndex, cell.getColumnIndex + 1)


        def upperCell():Cell =
            cell.getSheet.cell(cell.getRowIndex - 1, cell.getColumnIndex)

        def lowerCell():Cell =
            cell.getSheet.cell(cell.getRowIndex + 1, cell.getColumnIndex)

        def leftCell():Cell =
            cell.getSheet.cell(cell.getRowIndex, cell.getColumnIndex - 1)

        def rightCell():Cell =
            cell.getSheet.cell(cell.getRowIndex, cell.getColumnIndex + 1)
        

        ////////////////////////////////////////////////////////////////
        // Cell Stream (Reader)
        //
        def getUpperStream():Stream[Option[Cell]] = {
            def inner(sheet:Sheet, rownum:Int, colnum:Int):Stream[Option[Cell]] =
                Stream.cons(sheet.getCellOption(rownum, colnum),
                    if (rownum > 0)
                        inner(sheet, rownum - 1, colnum)
                    else
                        Stream.empty
                    )
            inner(cell.getSheet, cell.getRowIndex, cell.getColumnIndex)
        }

        def getLowerStream():Stream[Option[Cell]] = {
            def inner(sheet:Sheet, rownum:Int, colnum:Int):Stream[Option[Cell]] =
                Stream.cons(sheet.getCellOption(rownum, colnum), inner(sheet, rownum + 1, colnum))
            inner(cell.getSheet, cell.getRowIndex, cell.getColumnIndex)
        }

        def getLeftStream():Stream[Option[Cell]] = {
            def inner(sheet:Sheet, rownum:Int, colnum:Int):Stream[Option[Cell]] =
                Stream.cons(sheet.getCellOption(rownum, colnum),
                    if (colnum > 0)
                        inner(sheet, rownum, colnum - 1)
                    else
                        Stream.empty
                    )
            inner(cell.getSheet, cell.getRowIndex, cell.getColumnIndex)
        }

        def getRightStream():Stream[Option[Cell]] = {
            def inner(sheet:Sheet, rownum:Int, colnum:Int):Stream[Option[Cell]] =
                Stream.cons(sheet.getCellOption(rownum, colnum), inner(sheet, rownum, colnum + 1))
            inner(cell.getSheet, cell.getRowIndex, cell.getColumnIndex)
        }

        ////////////////////////////////////////////////////////////////
        // Cell Stream (Writer)
        //
        def upperStream():Stream[Cell] = {
            Stream.cons(cell, Try(cell.upperCell) match {
                case Success(next) => next.upperStream
                case Failure(e) => Stream.empty
            })
        }

        def lowerStream():Stream[Cell] = {
            Stream.cons(cell, Try(cell.lowerCell) match {
                case Success(next) => next.lowerStream
                case Failure(e) => Stream.empty
            })
        }

        def leftStream():Stream[Cell] = {
            Stream.cons(cell, Try(cell.leftCell) match {
                case Success(next) => next.leftStream
                case Failure(e) => Stream.empty
            })
        }

        def rightStream():Stream[Cell] = {
            Stream.cons(cell, Try(cell.rightCell) match {
                case Success(next) => next.rightStream
                case Failure(e) => Stream.empty
            })
        }

        ////////////////////////////////////////////////////////////////
        // CellStyle
        //
        def setBorderTop(borderStyle:BorderStyle):Unit = {
            val cellStyle = cell.getCellStyle
            val t = cellStyle.toTuple
            val newTuple = (borderStyle, t._2, t._3, t._4, t._5,
                t._6, t._7, t._8, t._9, t._10)
            val workbook = cell.getSheet.getWorkbook

            workbook.findCellStyle(newTuple) match {
                case Some(s) => cell.setCellStyle(s)
                case None => {
                    val newStyle = workbook.createCellStyle
                    newStyle.cloneStyleFrom(cellStyle)
                    newStyle.setBorderTop(borderStyle)
                    cell.setCellStyle(newStyle)
                }
            }
        }

        def setBorderBottom(borderStyle:BorderStyle):Unit = {
            val cellStyle = cell.getCellStyle
            val t = cellStyle.toTuple
            val newTuple = (t._1, borderStyle, t._3, t._4, t._5,
                t._6, t._7, t._8, t._9, t._10)
            val workbook = cell.getSheet.getWorkbook

            workbook.findCellStyle(newTuple) match {
                case Some(s) => cell.setCellStyle(s)
                case None => {
                    val newStyle = workbook.createCellStyle
                    newStyle.cloneStyleFrom(cellStyle)
                    newStyle.setBorderBottom(borderStyle)
                    cell.setCellStyle(newStyle)
                }
            }
        }

        def setBorderLeft(borderStyle:BorderStyle):Unit = {
            val cellStyle = cell.getCellStyle
            val t = cellStyle.toTuple
            val newTuple = (t._1, t._2, borderStyle, t._4, t._5,
                t._6, t._7, t._8, t._9, t._10)
            val workbook = cell.getSheet.getWorkbook

            workbook.findCellStyle(newTuple) match {
                case Some(s) => cell.setCellStyle(s)
                case None => {
                    val newStyle = workbook.createCellStyle
                    newStyle.cloneStyleFrom(cellStyle)
                    newStyle.setBorderLeft(borderStyle)
                    cell.setCellStyle(newStyle)
                }
            }
        }

        def setBorderRight(borderStyle:BorderStyle):Unit = {
            val cellStyle = cell.getCellStyle
            val t = cellStyle.toTuple
            val newTuple = (t._1, t._2, t._3, borderStyle,
                t._5, t._6, t._7, t._8, t._9, t._10)
            val workbook = cell.getSheet.getWorkbook

            workbook.findCellStyle(newTuple) match {
                case Some(s) => cell.setCellStyle(s)
                case None => {
                    val newStyle = workbook.createCellStyle
                    newStyle.cloneStyleFrom(cellStyle)
                    newStyle.setBorderRight(borderStyle)
                    cell.setCellStyle(newStyle)
                }
            }
        }
    }

    implicit class CellStyleImplicit(cellStyle:CellStyle) {

        def toTuple():(
            BorderStyle, 
            BorderStyle,
            BorderStyle,
            BorderStyle,
            Color,
            Color,
            FillPatternType,
            HorizontalAlignment,
            VerticalAlignment,
            Boolean) = (
                cellStyle.getBorderTopEnum,
                cellStyle.getBorderBottomEnum,
                cellStyle.getBorderLeftEnum,
                cellStyle.getBorderRightEnum,
                cellStyle.getFillForegroundColorColor,
                cellStyle.getFillBackgroundColorColor,
                cellStyle.getFillPatternEnum,
                cellStyle.getAlignmentEnum,
                cellStyle.getVerticalAlignmentEnum,
                cellStyle.getWrapText)
    }

    implicit def CellStyleImplicit_(cellStyle:CellStyle) =
        cellStyle.toTuple
}
