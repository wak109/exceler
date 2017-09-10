/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.cell

import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.excel.excellib.ImplicitConversions._

case class XlsRect( 
  val sheet:Sheet,
  override val row:Int,
  override val col:Int,
  override val height:Int,
  override val width:Int
) extends Rect[XlsRect] {

  def getValue() = this

  // TODO: XML 
  lazy val text = (for {
    cnum <- (col until col + width).toStream
    rnum <- (row until row + height).toStream
    value <- sheet.cell(rnum, cnum).getValueString.map(_.trim)
  } yield value).headOption.getOrElse("")
}

case class XlsCell(
  val xlsRect:XlsRect,
  override val row:Int,
  override val col:Int,
  override val height:Int,
  override val width:Int
) extends XmlCell {

  // TODO: tag <p> is OK??
  def getValue():Elem = <p>{xlsRect.text}</p>
}

