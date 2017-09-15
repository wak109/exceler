/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler

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
import exceler.tablex._

import CommonLib._


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
      book <- ExcelerBook.getBook(filename)
      sheet <- book.getSheetOption(sheetname)
      table <- XlsTable(sheet).get(tablename)
    } yield AbcTableQuery[XlsRect](table).queryByString(rowKeys, colKeys, blockKey)
  }
}


object ExcelerBook {

  lazy private val bookMap = 
    getListOfFiles(Config().excelDir)
      .filter(_.canRead)
      .filter(_.getName.endsWith(".xlsx"))
      .map((f:File)=>(f.getName, WorkbookFactory.create(f, null ,true)))
      .toMap

  def getBook(bookName:String) = bookMap.get(bookName)
}

