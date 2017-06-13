/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection.JavaConverters._
import scala.language.implicitConversions
import scala.util.control.Exception._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.xmlbeans.XmlObject

import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape


package ExcelLib {
    package Shape {
        trait ImplicitConversions {

            implicit class ToExcelShapeSheet(val sheet: Sheet)
              extends ExcelShapeSheetExtra

            implicit class ToExcelShapeXSSFShape(val shape: XSSFShape)
              extends ExcelShapeXSSFShapeExtra

        }
        object ImplicitConversions extends ImplicitConversions
    }
}

class ExcelSimpleShape(
    drawing:XSSFDrawing,
    ctShape:CTShape
    ) extends XSSFSimpleShape(drawing, ctShape)


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
