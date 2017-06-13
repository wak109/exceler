/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection.JavaConverters._
import scala.language.implicitConversions
import org.scalatest.FunSuite

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel._
import org.apache.xmlbeans.XmlObject
import org.apache.poi.POIXMLDocumentPart
import org.openxmlformats.schemas.drawingml.x2006.spreadsheetDrawing.CTShape


import java.io.File
import java.nio.file.{Paths, Files}

import ExcelLib._
import ExcelShape.Conversion._

class ExcelShapeTest extends FunSuite with TestCommon {

    test("Shapes") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val relations = workbook.getSheet("shapes").asInstanceOf[POIXMLDocumentPart].getRelations.asScala
        val drawing = (
            for {
                xfd <- relations
                if xfd.isInstanceOf[XSSFDrawing]
            } yield xfd.asInstanceOf[XSSFDrawing]
        ).headOption 
        drawing match {
            case Some(xd) => assert(xd.getShapes.asScala.length == 8)
            case None => assert(false)
        }
    }

    test("getShapes") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheet("shapes")
        assert(sheet.getXSSFShapes.length == 8)
    }

    test("XSSFShapeExt") {
        val file = new File(getClass.getResource(testWorkbook1).toURI)
        val workbook = WorkbookFactory.create(file)
        val sheet = workbook.getSheet("shapes")
        for (shape <- sheet.getXSSFShapes) {
           println(shape.toXmlObject)
        }
    }
}
