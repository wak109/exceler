/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.app

import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import org.scalatra._

import java.io._
import java.nio.file._

import exceler.common._
import exceler.excel._
import excellib.ImplicitConversions._
import exceler.abc.{AbcTableQuery}
import exceler.xls.{XlsRect,XlsTable}

import CommonLib._


class ExcelerServlet extends ScalatraServlet {

  get("/:book/:sheet/:table") {

    val result = Exceler.query(
      params("book"),
      params("sheet"),
      params("table"),
      params.getOrElse("row", ""),
      params.getOrElse("column", ""),
      params.getOrElse("block", "")
    )

    <html>
      <body>
        <p>Result: {result.fold("")(
            _.map((x)=>x.xml.text).reduce((a,b)=>a + "," + b))}</p>
      </body>
    </html>
  }

  get("/") {

    val bookList = Exceler.getBookList
    
    <books>
    { for (book <- bookList) yield <book name={book} /> }
    </books>
  }
}




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

