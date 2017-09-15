/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.app

import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

import exceler.excel.excellib.ImplicitConversions._

import exceler.common._
import exceler.abc.{AbcTableQuery}
import exceler.xls.{XlsRect,XlsTable}

import CommonLib._


case class ExcelerBook(val workbook:Workbook) {

}

object Exceler {

  def query(
    filename:String,
    sheetname:String,
    tablename:String,
    rowKeys:String,
    colKeys:String,
    blockKey:String)
  (implicit conv:(XlsRect=>String)) = {

    for {
      book <- getBook(filename)
      sheet <- book.getSheetOption(sheetname)
      table <- XlsTable(sheet).get(tablename)
    } yield AbcTableQuery[XlsRect](table).queryByString(rowKeys, colKeys, blockKey)
  }

  lazy private val bookMap = 
    getListOfFiles(Config().excelDir)
      .filter(_.canRead)
      .filter(_.getName.endsWith(".xlsx"))
      .map((f:File)=>(f.getName, ExcelerBook(
          WorkbookFactory.create(f, null ,true)))).toMap

  def getBookList() = bookMap.keys.toList
  def getBook(bookName:String) = bookMap.get(bookName).map(_.workbook)
}

