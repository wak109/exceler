/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xml

import scala.collection._
import scala.language.implicitConversions

trait XmlRect {
  val top:Int
  val left:Int
  val height:Int
  val width:Int
}

trait XmlCell extends XmlRect
