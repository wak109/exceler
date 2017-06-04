/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import ExcelRectangle._

object Exceler {

    def convertExcelTableToXML(filename:String):Try[Unit] = {

        Try {
            val file = new File(filename)
            val workbook = WorkbookFactory.create(file, null ,true) 
            for {
                sheet <- workbook.sheetIterator.asScala
                rect <- sheet.getRectangleList
                trow <- rect.getInnerRectangleList
                tcell <- trow
            } {
                println(tcell)
            }
        }
    }
}
