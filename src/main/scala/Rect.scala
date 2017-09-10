/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.rect

import scala.collection._
import scala.language.implicitConversions

trait Loc {
  val top:Int
  val left:Int
}

trait Rect extends Loc {
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

object Table {
  def apply[T<:Loc](locList:Seq[T]) = locList.groupBy(_.top)
    .toList.sortBy(_._1).map(_._2).map(_.sortBy(_.left))
}

case class Pad[T<:Rect](val top:Int, val left:Int, val rect:T)

case class Combo[T<:Rect](val combo:Either[T,Pad[T]]) extends Loc {

  override val top = combo match {
    case Left(rect) => rect.top
    case Right(pad) => pad.top
  }

  override val left = combo match {
    case Left(rect) => rect.left
    case Right(pad) => pad.left
  }
}

object Combo {
  implicit def comboToRect[T<:Rect](c:Combo[T]):T = c.combo match {
    case Left(rect) => rect
    case Right(pad) => pad.rect
  }
}
