/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection.JavaConverters._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._

import org.apache.xmlbeans.XmlObject

import java.io._
import java.nio.file._


object ExcelLib {

    implicit class ToWorkbookExtra(w:Workbook)
            extends Holder(w) with WorkbookExtra

    implicit class ToSheetExtra(s:Sheet)
            extends Holder(s) with SheetExtra

    implicit class ToRowExtra(r:Row)
            extends Holder(r) with RowExtra

    implicit class ToCellExtra(c:Cell)
            extends Holder(c) with CellExtra
            with CellBorderExtra
            with CellOuterBorderExtra

    implicit class ToCellStyleExtra(c:CellStyle)
            extends Holder(c) with CellStyleExtra

    implicit class ToXSSFShapeExtra(x:XSSFShape)
            extends Holder(x) with XSSFShapeExtra

    implicit def toCellStyleTuple(cellStyle:CellStyle) =
            cellStyle.toTuple
}


class Holder[T](val obj:T)

object Holder {
    implicit def toObj[T](holder:Holder[T]):T = holder.obj
}


trait WorkbookExtra {
    workbook:Holder[Workbook] =>

    import ExcelLib._    


    def saveAs(filename:String): Unit =  {
        FileLib.createParentDir(filename)
        val out = new FileOutputStream(filename)
        workbook.write(out)
        out.close()
    }

    def getSheetOption(name:String):Option[Sheet] = Option(
            workbook.getSheet(name))

    def sheet(name:String):Sheet = {
        this.getSheetOption(name) match {
            case Some(s) => s
            case None => Try(workbook.createSheet(name)) match {
                case Success(s) => s
                case Failure(e) => throw e
            }
        }
    }


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
    def findCellStyle(tuple:(
            BorderStyle, 
            BorderStyle,
            BorderStyle,
            BorderStyle,
            Color,
            Color,
            FillPatternType,
            HorizontalAlignment,
            VerticalAlignment,
            Boolean
            )):Option[CellStyle] = (
        for {
            i <- (0 until workbook.getNumCellStyles).toStream
            cellStyle = workbook.getCellStyleAt(i)
            if cellStyle.toTuple == tuple
        } yield cellStyle
    ).headOption
}


trait SheetExtra {
    sheet:Holder[Sheet] => 

    import ExcelLib._    


    def getRowOption(rownum:Int):Option[Row] =
            Option(sheet.getRow(rownum))

    def row(rownum:Int):Row = {
        this.getRowOption(rownum) match {
            case Some(r) => r
            case None => sheet.createRow(rownum)
        }
    }

    def getCellOption(rownum:Int, colnum:Int):Option[Cell] = 
            sheet.getRowOption(rownum).flatMap(_.getCellOption(colnum))

    def cell(rownum:Int, colnum:Int):Cell =
            sheet.row(rownum).cell(colnum)

    def getDrawingPatriarchOption():Option[Drawing[_ <: Shape]] =
            Option(sheet.getDrawingPatriarch)

    def drawingPatriarch():Drawing[_ <: Shape] = {
        this.getDrawingPatriarchOption match {
            case Some(d) => d
            case None => sheet.createDrawingPatriarch()
        }
    }

    def getXSSFShapes():List[XSSFShape] = {
        sheet.getDrawingPatriarchOption match {
            case Some(drawing) => drawing.asInstanceOf[
                    XSSFDrawing].getShapes.asScala.toList
            case None => List()
        }
    }
}


trait RowExtra {
    row:Holder[Row] =>

    import ExcelLib._    


    def getCellOption(colnum:Int):Option[Cell] =
        Option(row.getCell(colnum))

    def cell(colnum:Int):Cell =
        row.getCell(colnum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
}


trait CellExtra {
    cell:Holder[Cell] =>

    import ExcelLib._    

    def getValue_():Any = {
        cell.getCellTypeEnum match {
            case CellType.BLANK => cell.getStringCellValue
            case CellType.BOOLEAN => cell.getBooleanCellValue
            case CellType.ERROR => cell.getErrorCellValue
            case CellType.FORMULA =>
                cell.getCachedFormulaResultTypeEnum match {
                    case CellType.BLANK => cell.getStringCellValue
                    case CellType.BOOLEAN => cell.getBooleanCellValue
                    case CellType.ERROR => cell.getErrorCellValue
                    case CellType.NUMERIC => cell.getNumericCellValue
                    case CellType.STRING => cell.getStringCellValue
                    case _ => cell.getStringCellValue
                }
            case CellType.NUMERIC => cell.getNumericCellValue
            case CellType.STRING => cell.getStringCellValue
            case _ => cell.getStringCellValue
        }
    }

    def getValueString():Option[String] = {
        cell.getValue_ match {
            case null => None
            case v => v.toString match {
                case "" => None
                case s => Some(s)
            }
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
}


trait CellBorderExtra {
    cell:Holder[Cell] with CellExtra =>

    import ExcelLib._    


    ////////////////////////////////////////////////////////////
    // setBorder
    //
    def setBorderTop(borderStyle:BorderStyle):Unit = {
        val cellStyle = cell.getCellStyle
        val newTuple = cellStyle.toTuple.copy(_1 = borderStyle)
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
        val newTuple = cellStyle.toTuple.copy(_2 = borderStyle)
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
        val newTuple = cellStyle.toTuple.copy(_3 = borderStyle)
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
        val newTuple = cellStyle.toTuple.copy(_4 = borderStyle)
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

    ////////////////////////////////////////////////////////////
    // hasBorder
    //
    def hasBorderBottom():Boolean = {
        (cell.getCellStyle.getBorderBottomEnum != BorderStyle.NONE) ||
        (cell.getLowerCell
            .map(_.getCellStyle.getBorderTopEnum != BorderStyle.NONE)
            .getOrElse(false))
    }
    
    def hasBorderTop():Boolean = {
        (cell.getRowIndex == 0) ||
        (cell.getCellStyle.getBorderTopEnum != BorderStyle.NONE) ||
        (cell.getUpperCell
            .map(_.getCellStyle.getBorderBottomEnum != BorderStyle.NONE)
            .getOrElse(false))
    }
    
    def hasBorderRight():Boolean = {
        (cell.getCellStyle.getBorderRightEnum != BorderStyle.NONE) ||
        (cell.getRightCell
            .map(_.getCellStyle.getBorderLeftEnum != BorderStyle.NONE)
            .getOrElse(false))
    }
    
    def hasBorderLeft():Boolean = {
        (cell.getColumnIndex == 0) ||
        (cell.getCellStyle.getBorderLeftEnum != BorderStyle.NONE) ||
        (cell.getLeftCell
            .map(_.getCellStyle.getBorderRightEnum != BorderStyle.NONE)
            .getOrElse(false))
    }

    ////////////////////////////////////////////////////////////
    // ThickBorder
    //
    def hasThickBorderBottom():Boolean = {
        (cell.getCellStyle.getBorderBottomEnum == BorderStyle.THICK) ||
        (cell.getLowerCell
            .map(_.getCellStyle.getBorderTopEnum == BorderStyle.THICK)
            .getOrElse(false))
    }

    def hasThickBorderTop():Boolean = {
        (cell.getRowIndex == 0) ||
        (cell.getCellStyle.getBorderTopEnum == BorderStyle.THICK) ||
        (cell.getUpperCell
            .map(_.getCellStyle.getBorderBottomEnum == BorderStyle.THICK)
            .getOrElse(false))
    }

    def hasThickBorderRight():Boolean = {
        (cell.getCellStyle.getBorderRightEnum == BorderStyle.THICK) ||
        (cell.getRightCell
            .map(_.getCellStyle.getBorderLeftEnum == BorderStyle.THICK)
            .getOrElse(false))
    }

    def hasThickBorderLeft():Boolean = {
        (cell.getColumnIndex == 0) ||
        (cell.getCellStyle.getBorderLeftEnum != BorderStyle.THICK) ||
        (cell.getLeftCell
            .map(_.getCellStyle.getBorderRightEnum == BorderStyle.THICK)
            .getOrElse(false))
    }
}

trait CellOuterBorderExtra {
    cell:Holder[Cell] with CellExtra with CellBorderExtra =>

    import ExcelLib._

    ////////////////////////////////////////////////////////////
    // isOuterBorder
    //
    def isOuterBorderTop():Boolean = {
        val upperCell = cell.getUpperCell
    
        (cell.hasBorderTop) &&
        (upperCell match {
            case Some(cell) =>
                (! cell.hasBorderLeft) &&
                (! cell.hasBorderRight)
            case None => true
        })
    }
    
    def isOuterBorderBottom():Boolean = {
        val lowerCell = cell.getLowerCell
    
        (cell.hasBorderBottom) &&
        (lowerCell match {
            case Some(cell) =>
                (! cell.hasBorderLeft) &&
                (! cell.hasBorderRight)
            case None => true
        })
    }
    
    def isOuterBorderLeft():Boolean = {
        val leftCell = cell.getLeftCell
    
        (cell.hasBorderLeft) &&
        (leftCell match {
            case Some(cell) =>
                (! cell.hasBorderTop) &&
                (! cell.hasBorderBottom)
            case None => true
        })
    }
    
    def isOuterBorderRight():Boolean = {
        val rightCell = cell.getRightCell
    
        (cell.hasBorderRight) &&
        (rightCell match {
            case Some(cell) =>
                (! cell.hasBorderTop) &&
                (! cell.hasBorderBottom)
            case None => true
        })
    }
}


trait CellStyleExtra {
    cellStyle:Holder[CellStyle] =>

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


trait XSSFShapeExtra {
    shape:Holder[XSSFShape] =>

    def toXmlObject():XmlObject = {
        shape.obj match {
            case x:XSSFSimpleShape =>
                    x.asInstanceOf[XSSFSimpleShape].getCTShape
            case x:XSSFConnector =>
                    x.asInstanceOf[XSSFConnector].getCTConnector
            case x:XSSFGraphicFrame =>
                    x.asInstanceOf[XSSFGraphicFrame].getCTGraphicalObjectFrame
            case x:XSSFPicture =>
                    x.asInstanceOf[XSSFPicture].getCTPicture
            case x:XSSFShapeGroup =>
                    x.asInstanceOf[XSSFShapeGroup].getCTGroupShape
        }
    }
}
