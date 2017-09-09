/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xml

import scala.xml.Elem
import exceler.rect.Rect

trait XmlCell extends Rect {
  def toXml():Elem
}
