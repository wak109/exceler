/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.excel

import scala.collection.JavaConverters._
import scala.language.implicitConversions
import scala.util.control.Exception._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.poi.hssf.usermodel._
import org.apache.xmlbeans.XmlObject

import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.{
  CTShape
}
import org.openxmlformats.schemas.drawingml.x2006.main.{
  CTSolidColorFillProperties
}


package excellib.Shape {
  trait ImplicitConversions {
    implicit class ToExcelShapeSheet(val sheet: Sheet)
      extends ExcelShapeSheetExtra
    implicit class ToExcelShapeXSSFShapeExtra(val shape:XSSFShape)
      extends ExcelShapeXSSFShapeExtra
    implicit class ToXSSFSimpleShapeExtra(
      val shape: XSSFSimpleShape) extends ExcelSimpleShapeExtra
  }
  object ImplicitConversions extends ImplicitConversions

}

import excellib.Shape.ImplicitConversions._

class ExcelSimpleShape(
  drawing:XSSFDrawing,
  ctShape:CTShape
  ) extends XSSFSimpleShape(drawing, ctShape)

trait Helper {
  def byte3ToRGB(b3:Array[Byte]):List[Int] = {
    assert(b3.length == 3)
    b3.toList.map((b:Byte)=>(b.toInt + 256) % 256)
  }

  def solidColorToRGB(scfp:CTSolidColorFillProperties) = {
    if (scfp.isSetSrgbClr)
      byte3ToRGB(scfp.getSrgbClr.getVal)
    else if (scfp.isSetHslClr) {
      println("Hsl")
      List(0, 0, 0) // ToDo
    }
    else if (scfp.isSetPrstClr) {
      println("Prst")
      List(0, 0, 0) // ToDo
    }
    else if (scfp.isSetSchemeClr) {
      println("Scheme")
      List(0, 0, 0) // ToDo
    }
    else if (scfp.isSetSysClr) 
      byte3ToRGB(scfp.getSysClr.getLastClr)
    else if (scfp.isSetScrgbClr) {
      println("Scrgb")
      List(0, 0, 0) // ToDo
    }
    else {
      println("???")
      List(0, 0, 0) // ToDo
    }
  }
}

trait ExcelSimpleShapeExtra extends Helper {
  val shape:XSSFSimpleShape

  def copyFrom(from:XSSFSimpleShape): Unit = {
    shape.setBottomInset(from.getBottomInset)
    shape.setLeftInset(from.getLeftInset)
    shape.setRightInset(from.getRightInset)
    shape.setShapeType(from.getShapeType)
    shape.setText(from.getText)
    shape.setTextAutofit(from.getTextAutofit)
    shape.setTextDirection(from.getTextDirection)
    shape.setTextHorizontalOverflow(from.getTextHorizontalOverflow)
    shape.setTextVerticalOverflow(from.getTextVerticalOverflow)
    shape.setTopInset(from.getTopInset)
    shape.setVerticalAlignment(from.getVerticalAlignment)
    shape.setWordWrap(from.getWordWrap)

    shape.setLineWidth(from.getLineWidth)
    val lc = from.getLineStyleColor
    shape.setLineStyleColor(lc(0), lc(1), lc(2))
    val fc = from.getFillColor
    shape.setFillColor(fc(0), fc(1), fc(2))
  }

  // 12700 = 1 pt
  def getLineWidth() = {
    shape.getCTShape.getSpPr.getLn.getW.toDouble / 12700
  }

  def getLineStyleColor() = {
    solidColorToRGB(shape.getCTShape.getSpPr.getLn.getSolidFill)
  }

  def getFillColor() = {
    solidColorToRGB(shape.getCTShape.getSpPr.getSolidFill)
  }
}

trait ExcelShapeSheetExtra {
  val sheet:Sheet

  def getDrawingPatriarchOption():Option[XSSFDrawing] =
      Option(sheet.asInstanceOf[XSSFSheet].getDrawingPatriarch)

  def drawingPatriarch():XSSFDrawing = {
    this.getDrawingPatriarchOption match {
      case Some(d) => d
      case None => sheet.asInstanceOf[XSSFSheet].createDrawingPatriarch()
    }
  }

  def getXSSFShapes():List[XSSFShape] = {
    this.getDrawingPatriarchOption match {
      case Some(drawing) => drawing.asInstanceOf[
          XSSFDrawing].getShapes.asScala.toList
      case None => List()
    }
  }
}

trait ExcelShapeXSSFShapeExtra {
  val shape:XSSFShape

  def toXmlObject():XmlObject = {
    shape match {
      case x:XSSFSimpleShape =>
          x.asInstanceOf[XSSFSimpleShape].getCTShape
      case x:XSSFConnector =>
          x.asInstanceOf[XSSFConnector].getCTConnector
      case x:XSSFGraphicFrame =>
          x.asInstanceOf[XSSFGraphicFrame]
              .getCTGraphicalObjectFrame
      case x:XSSFPicture =>
          x.asInstanceOf[XSSFPicture].getCTPicture
      case x:XSSFShapeGroup =>
          x.asInstanceOf[XSSFShapeGroup].getCTGroupShape
    }
  }
}
