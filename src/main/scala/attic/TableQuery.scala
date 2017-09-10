/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.table

import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import exceler.common._
import CommonLib.ImplicitConversions._


trait TableQueryFunction[T] {
  val tableQueryFunction:QueryFunction

  trait QueryFunction {
    def create(rect:T):TableQuery[T]
  }
}


trait TableQuery[T] extends Table[T]
    with TableFunction[T] with TableQueryFunction[T] {
  rect:T =>

  def query(
    rowpredList:List[(String) => Boolean],
    colpredList:List[(String) => Boolean]
  ):List[List[T]] = {
    for {
      row <- this.queryRow(rowpredList)
    } yield for {
      col <- this.queryColumn(colpredList)
    } yield tableFunction.getCross(row, col)
  }

  def query(
    rowpred:(String) => Boolean,
    colpred:(String) => Boolean
  ):List[List[T]] = {
    for {
      row <- this.queryRow(rowpred)
    } yield for {
      col <- this.queryColumn(colpred)
    } yield tableFunction.getCross(row, col)
  }

  def queryRow(
    predList:List[String => Boolean]
  ):List[T] = {
    predList match {
      case Nil => List[T](rect)
      case pred::predTail => for {
        row <- this.queryRow(pred)
        row <- tableQueryFunction.create(row).queryRow(predTail)
      } yield row
    }
  }

  def queryRow(
    pred:String => Boolean
  ):List[T] = {
    for {
      row <- this.rowList
      (colHead, colTail) = tableFunction.getHeadCol(row)
      if pred(tableFunction.getValue(colHead).getOrElse(""))
      col <- colTail
    } yield col
  }

  def queryColumn(
      predList:List[String => Boolean]
  ):List[T] = {
    predList match {
      case Nil => List[T](rect)
      case pred::predTail => for {
        col <- this.queryColumn(pred)
        col <- tableQueryFunction.create(col).queryColumn(predTail)
      } yield col
    }
  }

  def queryColumn(
    pred:String => Boolean
  ) :List[T] = {
    for {
      col <- this.colList
      (rowHead, rowTail) = tableFunction.getHeadRow(col)
      if pred(tableFunction.getValue(rowHead).getOrElse(""))
      row <- rowTail
    } yield row
  }
}


trait StackedTableQuery[T] extends TableQuery[T] {
  rect:T =>

  val isSeparator:(T)=>Boolean = tableFunction.getHeadCol(_)._2 == None

  lazy val tableMap = this.rowList
    .splitBy(isSeparator)
    .pairingBy(x=>isSeparator(tableFunction.mergeRect(x)))
    .map(pair=>(tableFunction.getValue(
      tableFunction.mergeRect(pair._1)).getOrElse(""),
      tableFunction.mergeRect(pair._2)))
    .map(pair => (pair._1, tableQueryFunction.create(pair._2)))
    .toMap

  override def queryRow(predList:List[String => Boolean]):List[T] = {
    predList match {
      case Nil => List[T](rect)
      case pred::predTail => {
        (for {
          (key, table) <- tableMap
          if pred(key)
          row <- table.queryRow(predTail) 
        } yield row).toList match {
          case Nil => super.queryRow(predList)
          case x => x
        }
      }
    }
  }
}
