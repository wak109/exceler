/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions

trait CellX[+T] {
  val row:Int
  val col:Int
  val value:T
}


case class ProxyCellX[+T] (
  override val row:Int,
  override val col:Int,
  val delegate:CellX[T]
) extends CellX[T] {
  override val value = delegate.value
}


case class UnitedCellX[+T](
  val united:Either[ProxyCellX[T],RangeX[T]])
    extends CellX[T] {

  override val row = united.fold(_.row, _.row)
  override val col = united.fold(_.col, _.col)
  override val value = united.fold(_.value, _.value)
}


object UnitedCellX {
  def apply[T](rect:RangeX[T]) = new UnitedCellX(Right(rect))
  def apply[T](holder:ProxyCellX[T]) = new UnitedCellX(Left(holder))
}


trait RangeX[+T] extends CellX[T] {
  val height:Int
  val width:Int

  val top:Int = row
  val left:Int = col
  val bottom:Int = row + height - 1
  val right:Int = col + width - 1

  lazy val cellList = UnitedCellX(this) +: (for {
    rnum <- (0 until height)
    cnum <- (0 until width)
    holder = ProxyCellX(row+rnum,col+cnum,this)
        if (rnum != 0) || (cnum != 0)
  } yield UnitedCellX(holder))
}


object TableX {
  def apply[T<:CellX[_]](cellList:Seq[T]) = cellList.groupBy(_.row)
    .toList.sortBy(_._1).map(_._2).map(_.sortBy(_.col))

  def toCompact[T](table:Seq[Seq[UnitedCellX[T]]]):Seq[Seq[RangeX[T]]] =
    this.apply(for {
      row <- table
      cell <- row
      rect <- cell.united.fold(_ => None, Some(_))
    } yield rect)

  def toArray[T](table:Seq[Seq[RangeX[T]]]):Seq[Seq[UnitedCellX[T]]] = 
    this.apply((for {
      row <- table
      cell <- row
    } yield cell.cellList).reduce((a, b)=> a ++ b))
}
