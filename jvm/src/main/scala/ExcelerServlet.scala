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
import java.util.Properties

import javax.servlet.http.HttpServlet

import exceler.common._
import exceler.excel._
import excellib.ImplicitConversions._
import exceler.abc.{AbcTableQuery}
import exceler.xls.{XlsRect,XlsTable}

import CommonLib._

class ExcelerConfig(servlet:HttpServlet) {

  private val properties = new Properties
  private val configFile = servlet.getInitParameter("configFile")

  if (Files.exists(Paths.get(configFile))) {
    val inputStream = new FileInputStream(configFile)
    properties.load(inputStream);
    inputStream.close();
  }
  val keySet = properties.keySet

  val excelDir = if (keySet.contains("excelDir"))
      properties.getProperty("excelDir")
    else
      servlet.getInitParameter("excelDir")
}

class ExcelerServlet extends ScalatraServlet {

  lazy val excelerConfig = new ExcelerConfig(this)
  lazy val exceler = new Exceler(excelerConfig.excelDir)

  get("/:book/:sheet/:table") {

    val result = this.exceler.query(
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

    val bookList = this.exceler.getBookList
    
    <books>
    { for (book <- bookList) yield <book name={book} /> }
    </books>
  }
}




case class ExcelerBook(val workbook:Workbook) {

}

class Exceler(excelDir:String) {

  private val bookMap = 
    getListOfFiles(this.excelDir)
      .filter(_.canRead)
      .filter(_.getName.endsWith(".xlsx"))
      .map((f:File)=>(f.getName, ExcelerBook(
          WorkbookFactory.create(f, null ,true)))).toMap

  def getBookList() = bookMap.keys.toList
  def getBook(bookName:String) = bookMap.get(bookName).map(_.workbook)

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

