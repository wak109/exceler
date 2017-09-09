/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xls

import scala.collection._
import scala.language.implicitConversions

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import exceler.common._
import exceler.xml._
import exceler.excel._

import CommonLib.ImplicitConversions._
import excellib.ImplicitConversions._
import excellib.Rectangle.ImplicitConversions._


case class XlsRect( 
  val sheet:Sheet,
  val top:Int,
  val left:Int,
  val height:Int,
  val width:Int
)

case class XlsCell(
  val xlsRect:XlsRect,
  override val top:Int,
  override val left:Int,
  override val height:Int,
  override val width:Int
) extends XmlCell {
  def this(xlsRect:XlsRect, xmlRect:(Int,Int,Int,Int)) =
    this(xlsRect, xmlRect._1, xmlRect._2, xmlRect._3, xmlRect._4)
}
