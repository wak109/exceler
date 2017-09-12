/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions
import scala.xml.Elem

import exceler.common.CommonLib.ImplicitConversions._

class QueryTableX[T](compactTable:Seq[Seq[RangeX[T]]])(
    implicit conv:(T=>String)) {  

  val arrayTable:Seq[Seq[UnitedCellX[T]]] = TableX.toArray(compactTable)

  val blockMap:Map[String,List[Int]] = (Range(0,compactTable.length)
    .toList.filter(r=>isSeparator(compactTable(r))):+compactTable.length)
    .pairwise.map(t=>((getCellString(t._1,0),
      Range(t._1+1,t._2).toList))).toMap

  /**
   * Using implicit conversion
   */
  def getCellString(row:Int,col:Int):String = arrayTable(row)(col).getValue

  def isSeparator(row:Seq[RangeX[T]]):Boolean = row.length == 1

  def query(
    blockKey:Option[String=>Boolean] = None,
    rowKeys:List[String=>Boolean] = Nil,
    colKeys:List[String=>Boolean] = Nil
  ):List[T] = {
    val block = queryBlockKey(blockKey)
    val rowList = queryRowKeys(rowKeys, 0, block)
    val colList = queryColKeys(colKeys, 0, Range(0,arrayTable(0).length).toList)
    for {
      row <- rowList
      col <- colList
    } yield arrayTable(row)(col).getValue
  }

  def queryBlockKey(blockKey:Option[String=>Boolean]):List[Int] = {
    blockKey match {
      case None => Range(0, arrayTable.length).toList
      case Some(key) => blockMap.keys.filter(key)
                          .map(blockMap(_)).toList.flatten
    }
  }

  def queryRowKeys(rowKeys:List[String=>Boolean], col:Int,
        rowList:List[Int]):List[Int] = rowKeys match {
    case Nil => rowList
    case head::tail => queryRowKeys(tail, col +1,
      rowList.filter(row=>head(getCellString(row, col))))
  }

  def queryColKeys(colKeys:List[String=>Boolean], row:Int,
        colList:List[Int]):List[Int] = colKeys match {
    case Nil => colList
    case head::tail => queryColKeys(tail, row +1,
      colList.filter(col=>head(getCellString(row, col))))
  }
}
