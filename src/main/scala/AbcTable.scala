/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions

trait AbcCell[+T] {
  val row:Int
  val col:Int
  val value:T
}


case class PlaceHolder[+T] (
  override val row:Int,
  override val col:Int,
  val delegate:AbcCell[T]
) extends AbcCell[T] {
  override val value = delegate.value
}


case class GenCell[+T](
  val united:Either[PlaceHolder[T],AbcRange[T]])
    extends AbcCell[T] {

  override val row = united.fold(_.row, _.row)
  override val col = united.fold(_.col, _.col)
  override val value = united.fold(_.value, _.value)
}


object GenCell {
  def apply[T](rect:AbcRange[T]) = new GenCell(Right(rect))
  def apply[T](holder:PlaceHolder[T]) = new GenCell(Left(holder))
}


trait AbcRange[+T] extends AbcCell[T] {
  val height:Int
  val width:Int

  val top:Int = row
  val left:Int = col
  val bottom:Int = row + height - 1
  val right:Int = col + width - 1

  lazy val cellList = GenCell(this) +: (for {
    rnum <- (0 until height)
    cnum <- (0 until width)
    holder = PlaceHolder(row+rnum,col+cnum,this)
        if (rnum != 0) || (cnum != 0)
  } yield GenCell(holder))
}


object AbcTable {
  def apply[T<:AbcCell[_]](cellList:Seq[T]) = cellList.groupBy(_.row)
    .toList.sortBy(_._1).map(_._2).map(_.sortBy(_.col))

  def toCompact[T](table:Seq[Seq[GenCell[T]]]):Seq[Seq[AbcRange[T]]] =
    this.apply(for {
      row <- table
      cell <- row
      rect <- cell.united.fold(_ => None, Some(_))
    } yield rect)

  def toArray[T](table:Seq[Seq[AbcRange[T]]]):Seq[Seq[GenCell[T]]] = 
    this.apply((for {
      row <- table
      cell <- row
    } yield cell.cellList).reduce((a, b)=> a ++ b))
}
