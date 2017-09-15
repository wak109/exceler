/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.table

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import exceler.common._
import CommonLib.ImplicitConversions._


trait TableFunction[T] {
  val tableFunction:Function

  trait Function {
    def getCross(row:T, col:T):T
    def getHeadRow(rect:T):(T,Option[T])
    def getHeadCol(rect:T):(T,Option[T])
    def getValue(rect:T):Option[String]
    def mergeRect(rectL:List[T]):T
  }
}


trait Table[T] extends TableFunction[T] {
  rect:T =>

  lazy val rowList:List[T] = this.getRowList(Some(rect))
  lazy val colList:List[T] = this.getColList(Some(rect))

  def getRowList(rect:Option[T]):List[T] = {
    rect match {
      case None => Nil
      case Some(rect) => {
        val (headRow, tailRow) = tableFunction.getHeadRow(rect)
        headRow :: getRowList(tailRow)
      }
    }
  }

  def getColList(rect:Option[T]):List[T] = {
    rect match {
      case None => Nil
      case Some(rect) => {
        val (headCol, tailCol) = tableFunction.getHeadCol(rect)
        headCol :: getColList(tailCol)
      }
    }
  }
}
