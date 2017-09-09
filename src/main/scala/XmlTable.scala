/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.xml

import scala.collection._
import scala.language.implicitConversions

object XmlTable {

  def apply(xmlCellList:Seq[XmlCell]) = xmlCellList.groupBy(_.top)
    .toList.sortBy(_._1).map(_._2).map(_.sortBy(_.left))
}
