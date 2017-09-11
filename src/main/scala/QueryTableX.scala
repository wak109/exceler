/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions

trait QueryTableX[T] {
  val compactTable:Seq[Seq[RangeX[T]]]  
  val arrayTable:Seq[Seq[UnitedCellX[T]]]  



  private def isSeparator(row:Seq[RangeX[T]]):Boolean = row.length == 1
}
