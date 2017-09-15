/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

import scala.collection._
import scala.language.implicitConversions
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io.File
import java.nio.file.{Paths, Files}

import org.scalatest.FunSuite

import exceler.tablex._
import exceler.common.CommonLib.ImplicitConversions._

class AbcTableQueryTest extends FunSuite with TestResource {

  import AbcTableQuery._

  implicit def elemToString(elem:Elem):String = elem.text

  val file = new File(getURI(testWorkbook1))
  val workbook = WorkbookFactory.create(file)
  val sheet = workbook.getSheet("stack")
  val compactTable = XlsTable(sheet,1,1,18,12)
  val qTable = new AbcTableQuery[XlsRect](XlsTable(sheet,1,1,18,12))

  val compactTable2 = XlsTable(sheet,1,1,18,12)
  val qTable2 = new AbcTableQuery[XlsRect](XlsTable(sheet,1,1,18,12))


  /*
  test("blockMap") {
    assert(qTable.blockMap("separator1") == Range(6,11).toList)
    assert(qTable.blockMap("separator2") == Range(12,17).toList)
  }
  */

  test("queryBlock") {
    assert(qTable.queryBlockKey(genKey("separator1"))
            == Range(6,11).toList)
    assert(qTable.queryBlockKey(genKey("separator2"))
            == Range(12,17).toList)
    assert(qTable.queryBlockKey(None)
            == Range(0,17).toList)
  }

  test("queryRowKeys") {
    assert(qTable.queryRowKeys(genKeyList("row1"),0,Range(0,17).toList)
      == List(1,6,12))
    assert(qTable.queryRowKeys(genKeyList("hehe"),0,Range(0,17).toList)
      == Nil)
  }

  test("queryColKeys") {
    assert(qTable.queryColKeys(genKeyList("col2"),0,Range(0,3).toList)
      == List(1))
    assert(qTable.queryColKeys(genKeyList("hehe"),0,Range(0,3).toList)
      == Nil)
  }

  test("query(Elem)") {
    assert(qTable.queryByString("row1", "col2").map(_.xml.text)
      == List("val12", "val12-1", "val12-2"))
    assert(qTable.queryByString("row1", "col2", "separator1").map(_.xml.text)
      == List("val12-1"))
    assert(qTable.queryByString("row1", "col2", "separator2").map(_.xml.text)
      == List("val12-2"))
  }

  test("query(XlsRect)") {
    assert(qTable2.queryByString("row1", "col2").map(_.xml.text)
      == List("val12", "val12-1", "val12-2"))
    assert(qTable2.queryByString("row1", "col2", "separator1").map(_.xml.text)
      == List("val12-1"))

    val xlsRect = qTable2.queryByString("row1", "col2", "separator2")

    assert(xlsRect(0).row == 14 && xlsRect(0).col == 5 &&
           xlsRect(0).height == 1 && xlsRect(0).width == 4)
  }
}
