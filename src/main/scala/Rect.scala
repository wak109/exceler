/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.rect

import scala.collection._
import scala.language.implicitConversions

trait Rect {
  val top:Int
  val left:Int
  val height:Int
  val width:Int
}

object Rect {

  implicit def apply(_top:Int,_left:Int,_height:Int,_width:Int):Rect =
    new Rect {
      override val top = _top
      override val left = _left
      override val height = _height
      override val width = _width
    }

  implicit def apply(t:(Int,Int,Int,Int)):Rect =
    new Rect {
      override val top = t._1
      override val left = t._2
      override val height = t._3
      override val width = t._4
    }
}
