/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection.JavaConverters._
import scala.language.implicitConversions
import scala.util.control.Exception._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.xmlbeans.XmlObject

import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape


object ExcelShape {
    object Conversion {
        implicit class ToExcelShapeSheet(val sheet:Sheet)
            extends ExcelShapeSheetExtra

        implicit class ToExcelShapeXSSFShape(val shape:XSSFShape)
            extends ExcelShapeXSSFShapeExtra
    }
}

class ExcelSimpleShape(
    drawing:XSSFDrawing,
    ctShape:CTShape
    ) extends XSSFSimpleShape(drawing, ctShape)


trait ExcelShapeSheetExtra {
    val sheet:Sheet

    def getDrawingPatriarchOption():Option[Drawing[_ <: Shape]] =
            Option(sheet.getDrawingPatriarch)

    def drawingPatriarch():Drawing[_ <: Shape] = {
        this.getDrawingPatriarchOption match {
            case Some(d) => d
            case None => sheet.createDrawingPatriarch()
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
