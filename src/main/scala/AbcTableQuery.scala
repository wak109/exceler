/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.abc

import scala.collection._
import scala.language.implicitConversions
import scala.xml.Elem

import exceler.common.CommonLib.ImplicitConversions._

case class AbcTableQuery[T](compactTable:Seq[Seq[AbcRange[T]]]) {  

  import AbcTableQuery._

  val arrayTable:Seq[Seq[GenCell[T]]] = AbcTable.toArray(compactTable)

  val blockList:List[(GenCell[T],List[Int])] =
    (Range(0,compactTable.length).toList.filter(
      r=>isSeparator(compactTable(r))):+compactTable.length)
      .pairwise.map(t=>((arrayTable(t._1)(0),
      Range(t._1+1,t._2).toList))).toList

  def isSeparator(row:Seq[AbcRange[T]]):Boolean = row.length == 1

  def query(
    rowKeys:List[T=>Boolean] = Nil,
    colKeys:List[T=>Boolean] = Nil,
    blockKey:Option[T=>Boolean] = None
  ):List[T] = {
    val block = queryBlockKey(blockKey)
    val rowList = queryRowKeys(rowKeys, 0, block)
    val colList = queryColKeys(colKeys, 0,
          Range(0,arrayTable(0).length).toList)
    for {
      row <- rowList
      col <- colList
    } yield arrayTable(row)(col).value
  }

  def queryByString(
    rowKeys:String = "",
    colKeys:String = "",
    blockKey:String = ""
  )(implicit getString:(T=>String)):List[T] = query(
      genKeyList[T](rowKeys),
      genKeyList[T](colKeys),
      genKey[T](blockKey))

  def queryBlockKey(blockKey:Option[T=>Boolean]):List[Int] = {
    blockKey match {
      case None => Range(0, arrayTable.length).toList
      case Some(key) => {
        
        blockList.filter((t=>key(t._1.value))).map(_._2).toList.flatten
      }
    }
  }

  def queryRowKeys(rowKeys:List[T=>Boolean], col:Int,
        rowList:List[Int]):List[Int] = rowKeys match {
    case Nil => rowList
    case head::tail => queryRowKeys(tail, col +1,
      rowList.filter(row=>head(arrayTable(row)(col).value)))
  }

  def queryColKeys(colKeys:List[T=>Boolean], row:Int,
        colList:List[Int]):List[Int] = colKeys match {
    case Nil => colList
    case head::tail => queryColKeys(tail, row +1,
      colList.filter(col=>head(arrayTable(row)(col).value)))
  }
}

object AbcTableQuery {

  def genKeyList[T](key:String)(
        implicit getString:(T=>String)):List[T=>Boolean] =
          key.split(",").toList.map(genKey[T](_))
            .map(_.getOrElse((t:T)=>true))

  def genKey[T](key:String)(
        implicit getString:(T=>String)):Option[T=>Boolean] = key match {
    case "" => None
    case _ => Some((t:T)=>getString(t) == key)
  }
}
