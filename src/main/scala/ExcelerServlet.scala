import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import org.scalatra._

import java.io._
import java.nio.file._

import ExcelLib.ImplicitConversions._
import ExcelLib.Rectangle.ImplicitConversions._
import ExcelLib.Table.ImplicitConversions._


class ExcelerServlet extends ScalatraServlet with ExcelTableFunction {

  get("/query/:book/:sheet/:table") {
    
    def isSameStr(s:String): String => Boolean = {
        s match {
            case "" => (x:String) => true
            case _  => (x:String) => x == s
        }
    }
    /*
    val file = new File(params("book"))
    val workbook = WorkbookFactory.create(file, null ,true) 
    val result = for {
        sheet <- workbook.getSheetOption(params("sheet")).toSeq
        table <- sheet.getTableMap[TableQueryImpl].get(
                  params("table")).toSeq
        row <- table.query(
            params.getOrElse("row","").split(",").toList.map(isSameStr),
            params.getOrElse("column","").split(",").toList.map(isSameStr))
        cell <- row
        value <- tableFunction.getValue(cell)
    } yield value
    */

    val result = for {
      book <- ExcelerBook(params("book")).toSeq
      sheet <- book.getSheet(params("sheet")).toSeq
      table <- sheet.getTable(params("table")).toSeq
      row <- table.query(
        params.getOrElse("row","").split(",").toList.map(isSameStr),
        params.getOrElse("column","").split(",").toList.map(isSameStr))
      cell <- row
      value <- tableFunction.getValue(cell)
    } yield value

    <html>
      <body>
        <p>Result: {result}</p>
      </body>
    </html>
  }

  get("/") {
    <html>
      <body>
        <h1>Hello, world!</h1>
        <p>Scalatra</p>
      </body>
    </html>
  }
}
