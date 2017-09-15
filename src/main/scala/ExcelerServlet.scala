/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import org.scalatra._

import java.io._
import java.nio.file._

import exceler.Exceler
import exceler.excel._
import exceler.tablex._

import excellib.ImplicitConversions._


class ExcelerServlet extends ScalatraServlet {

  get("/:book/:sheet/:table") {

    val result = Exceler.query[XlsRect](
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
    <html>
      <body>
        <h1>Hello, world!</h1>
        <p>Scalatra</p>
      </body>
    </html>
  }
}
