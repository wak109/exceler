/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.collection.convert.ImplicitConversionsToScala._
import scala.language.implicitConversions
import scala.math
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util._

import java.io._
import java.nio.file._

import CommonLib._

package ExcelLib {

  trait ImplicitConversions {
    implicit class ToWorkbookExtra(val workbook:Workbook)
        extends WorkbookExtra

    implicit class ToSheetExtra(val sheet:Sheet) extends SheetExtra

    implicit class ToRowExtra(val row:Row) extends RowExtra

    implicit class ToCellExtra(val cell:Cell) extends CellExtra
        with CellBorderExtra
        with CellOuterBorderExtra

    implicit class ToCellStyleExtra(val cellStyle:CellStyle)
        extends CellStyleExtra

    implicit def toCellStyleTuple(cellStyle:CellStyle) =
        cellStyle.toTuple
  }

  object ImplicitConversions extends ImplicitConversions
}


import ExcelLib.ImplicitConversions._

////////////////////////////////////////////////////////////////////////
// WorkbookExtra

trait WorkbookExtra {
  val workbook:Workbook

  def saveAs(filename:String): Unit =  {
    createParentDir(filename)
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

  def removeSheet(name:String):Unit = {
    for {
      sheet <- this.getSheetOption(name)
    } {
      workbook.removeSheetAt(workbook.getSheetIndex(sheet))
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
  val sheet:Sheet

  def getRowOption(rownum:Int):Option[Row] =
      Option(sheet.getRow(rownum))

  def row(rownum:Int):Row = {
    this.getRowOption(rownum) match {
      case Some(r) => r
      case None => sheet.createRow(rownum)
    }
  }

  def getCellOption(rownum:Int, colnum:Int):Option[Cell] = 
      this.getRowOption(rownum).flatMap(_.getCellOption(colnum))

  def cell(rownum:Int, colnum:Int):Cell =
      this.row(rownum).cell(colnum)
}


trait RowExtra {
  val row:Row

  def getCellOption(colnum:Int):Option[Cell] =
    Option(row.getCell(colnum))

  def cell(colnum:Int):Cell =
    row.getCell(colnum, Row.MissingCellPolicy.CREATE_NULL_AS_BLANK)
}


trait CellExtra {
  val cell:Cell

  def doubleTo(d:Double):Any = 
    if (d == math.rint(d)) math.round(d) else d

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
          case CellType.NUMERIC =>
              doubleTo(cell.getNumericCellValue)
          case CellType.STRING => cell.getStringCellValue
          case _ => cell.getStringCellValue
        }
      case CellType.NUMERIC => doubleTo(cell.getNumericCellValue)
      case CellType.STRING => cell.getStringCellValue
      case _ => cell.getStringCellValue
    }
  }

  def getValueString():Option[String] = {
    this.getValue_ match {
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

  ////////////////////////////////////////////////////////////////
  // Merged Region
  //
  def getMergedRegion():Option[CellRangeAddress] = {
    (for {
      region <- cell.getSheet.getMergedRegions
      if region.isInRange(cell)
    } yield region).headOption
  }
}


trait CellBorderExtra {
  val cell:Cell 

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
    cell.getMergedRegion match {
      case Some(region)
        if cell.getRowIndex != region.getLastRow => false
      case _ =>
        (cell.getCellStyle.getBorderBottomEnum
          != BorderStyle.NONE) ||
        (cell.getLowerCell.map(_.getCellStyle.getBorderTopEnum
          != BorderStyle.NONE).getOrElse(false))
    }
  }
  
  def hasBorderTop():Boolean = {
    cell.getMergedRegion match {
      case Some(region)
        if cell.getRowIndex != region.getFirstRow => false
      case _ =>
        (cell.getRowIndex == 0) ||
        (cell.getCellStyle.getBorderTopEnum
          != BorderStyle.NONE) ||
        (cell.getUpperCell.map(
          _.getCellStyle.getBorderBottomEnum != BorderStyle.NONE)
          .getOrElse(false))
    }
  }
  
  def hasBorderRight():Boolean = {
    cell.getMergedRegion match {
      case Some(region)
        if cell.getColumnIndex != region.getLastColumn => false
      case _ =>
        (cell.getCellStyle.getBorderRightEnum
          != BorderStyle.NONE) ||
        (cell.getRightCell.map(
          _.getCellStyle.getBorderLeftEnum != BorderStyle.NONE)
          .getOrElse(false))
    }
  }
  
  def hasBorderLeft():Boolean = {
    cell.getMergedRegion match {
      case Some(region)
        if cell.getColumnIndex != region.getFirstColumn => false
      case _ =>
        (cell.getColumnIndex == 0) ||
        (cell.getCellStyle.getBorderLeftEnum
          != BorderStyle.NONE) ||
        (cell.getLeftCell.map(
          _.getCellStyle.getBorderRightEnum != BorderStyle.NONE)
          .getOrElse(false))
    }
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
  val cell:Cell

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
  val cellStyle:CellStyle

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
