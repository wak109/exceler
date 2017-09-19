/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xls

import scala.language.implicitConversions
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.abc.{AbcRange,AbcTable}
import exceler.excel.excellib.ImplicitConversions._

case class XlsRect( 
  val sheet:Sheet,
  override val row:Int,
  override val col:Int,
  override val height:Int,
  override val width:Int
) extends AbcRange[XlsRect] {

  override val value = this

  // TODO
  lazy val xml:Elem = <p>{(for {
      cnum <- (col until col + width).toStream
      rnum <- (row until row + height).toStream
      value <- sheet.cell(rnum, cnum).getValueString.map(_.trim)
  } yield value).headOption.getOrElse("")}</p>
}

object XlsRect {
  implicit def convToElem(rect:XlsRect):Elem = rect.xml
  implicit def convToString(rect:XlsRect):String = rect.xml.text
}

case class XlsCell(
  override val value:XlsRect,
  override val row:Int,
  override val col:Int,
  override val height:Int,
  override val width:Int
  ) extends AbcRange[XlsRect]
