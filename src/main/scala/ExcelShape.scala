/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.language.implicitConversions
import scala.util.control.Exception._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.xmlbeans.XmlObject

import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape


object ExcelShape {
    object Conversion 
        extends XSSFShapeConversion
}


trait XSSFShapeConversion {
    implicit class ToXSSFShapeExtra(val shape:XSSFShape)
            extends XSSFShapeExtra
}


class ExcelSimpleShape(
    drawing:XSSFDrawing,
    ctShape:CTShape
    ) extends XSSFSimpleShape(drawing, ctShape)


trait XSSFShapeExtra {
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
