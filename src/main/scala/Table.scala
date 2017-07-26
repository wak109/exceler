/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.collection._
import scala.language.implicitConversions
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import CommonLib.ImplicitConversions._


trait TableFunction[T] {
    def getCross(row:T, col:T):T
    def getHeadRow(rect:T):(T,Option[T])
    def getHeadCol(rect:T):(T,Option[T])
    def getValue(rect:T):Option[String]
    def getTableName(rect:T):(Option[String], T)
    def mergeRect(rectL:List[T]):T
}


trait Table[T] {
    rect:T =>
    val tableFunction:TableFunction[T]

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

trait TableQuery[T] extends Table[T] {
    rect:T =>
    val tableFunction:TableFunction[T]
    val createTableQuery:(T) => TableQuery[T]

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
                row <- createTableQuery(row).queryRow(predTail)
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
                col <- createTableQuery(col).queryColumn(predTail)
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

    val tableFunction:TableFunction[T]
    val createTableQuery:(T) => TableQuery[T]
    val isSeparator:(T)=>Boolean = tableFunction.getHeadCol(_)._2 == None

    lazy val tableMap = this.rowList
        .splitBy(isSeparator)
        .pairingBy(x=>isSeparator(tableFunction.mergeRect(x)))
        .map(pair=>(tableFunction.getValue(
            tableFunction.mergeRect(pair._1)).getOrElse(""),
            tableFunction.mergeRect(pair._2)))
        .map(pair => (pair._1, createTableQuery(pair._2)))
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