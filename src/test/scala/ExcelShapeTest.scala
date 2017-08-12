/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
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

import ExcelLib.ImplicitConversions._
import ExcelLib.Shape.ImplicitConversions._

class ExcelShapeTest extends FunSuite with ExcelLibResource {

  test("Sheet.getDrawingPatriarch") {
    val sheet = (new XSSFWorkbook()).sheet("test")

    sheet.getDrawingPatriarchOption match {
      case Some(c) => assert(false)
      case None => assert(true)
    }

    val drawing = sheet.drawingPatriarch

    println(drawing.getClass)


    sheet.getDrawingPatriarchOption match {
      case Some(c) => assert(true)
      case None => assert(false)
    }
  }

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

  test("print Xml of Shape") {
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("shapes")
    for (shape <- sheet.getXSSFShapes) {
       // println(shape.toXmlObject)
    }
  }

  test("copy shape") {
    
    val file = new File(getClass.getResource(testWorkbook1).toURI)
    val workbook = WorkbookFactory.create(file)
    val sheet = workbook.getSheet("shapes")

    val workbook2 = new XSSFWorkbook
    val sheet2 = workbook2.sheet("shapes")
    for {
      row <- (0 until 100)
      col <- (0 until 100)
    } sheet2.cell(row, col)

    val drawing2 = sheet2.drawingPatriarch

    for (shape <- sheet.getXSSFShapes) {
      shape match {
        case x:XSSFSimpleShape =>
          val anchor = shape.getAnchor.asInstanceOf[XSSFClientAnchor]
          val s = drawing2.createSimpleShape(anchor)
          s.asInstanceOf[XSSFSimpleShape].copyFrom(shape.asInstanceOf[XSSFSimpleShape])
        case _ =>
      }
    }

    workbook2.saveAs("hello.xlsx")
  }
}
