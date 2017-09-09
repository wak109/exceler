/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xls

import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.excel.excellib.ImplicitConversions._
import exceler.xml.XmlCell
import exceler.rect.Rect

case class XlsRect( 
  val sheet:Sheet,
  override val top:Int,
  override val left:Int,
  override val height:Int,
  override val width:Int
) extends Rect {

  def this(sheet:Sheet, rect:Rect) =
    this(sheet, rect.top, rect.left, rect.height, rect.width)

  // TODO: XML 
  def getString():String = (for {
    col <- (left until left + width).toStream
    row <- (top until top + height).toStream
    value <- sheet.cell(row, col).getValueString.map(_.trim)
  } yield value).headOption.getOrElse("")
}

case class XlsCell(
  val xlsRect:XlsRect,
  override val top:Int,
  override val left:Int,
  override val height:Int,
  override val width:Int
) extends XmlCell {

  def this(xlsRect:XlsRect, rect:Rect) =
    this(xlsRect, rect.top, rect.left, rect.height, rect.width)

  // TODO: tag <p> is OK??
  def toXml():Elem = <p>{xlsRect.getString}</p>
}

