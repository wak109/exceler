/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.tablex

import scala.collection._
import scala.language.implicitConversions
import scala.xml.Elem

import exceler.common.CommonLib.ImplicitConversions._

class QueryTableX(compactTable:Seq[Seq[RangeX[Elem]]]) {  

  val arrayTable:Seq[Seq[UnitedCellX[Elem]]] = TableX.toArray(compactTable)

  val blockMap = (Range(0,compactTable.length).toList
    .filter(r=>isSeparator(compactTable(r))):+compactTable.length)
    .pairwise.map(t=>((compactTable(t._1)(0).getValue.text,
      Range(t._1+1,t._2).toList))).toMap

  def isSeparator(row:Seq[RangeX[Elem]]):Boolean = row.length == 1

  def query(
    blockKey:Option[String=>Boolean] = None,
    rowKeys:List[String=>Boolean] = Nil,
    colKeys:List[String=>Boolean] = Nil
  ):List[Elem] = {
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
      rowList.filter(row=>head(arrayTable(row)(col).getValue.text)))
  }

  def queryColKeys(colKeys:List[String=>Boolean], row:Int,
        colList:List[Int]):List[Int] = colKeys match {
    case Nil => colList
    case head::tail => queryColKeys(tail, row +1,
      colList.filter(col=>head(arrayTable(row)(col).getValue.text)))
  }
}
