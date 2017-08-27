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

  get("/:book/:sheet/:table") {

    val result = TableQueryImpl.queryExcelTable(
      params("book"),
      params("sheet"),
      params("table"),
      params.getOrElse("row",""),
      params.getOrElse("column","")
    )

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
