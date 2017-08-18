/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import java.io._
import java.nio.file._

object CommonLib {

  object ImplicitConversions {
    implicit class ImplicitListConversion[T](val lst:List[T]) {
      def splitBy(pred:T=>Boolean) = splitListBy(pred, lst)
      def pairingBy(pred:T=>Boolean) = pairingListBy(pred, lst)
      def blockingBy(pred:T=>Boolean) = blockingListBy(pred, lst)
      def pairwise() = pairwiseBy(lst)
    }
  }

  def createParentDir(filename:String) : Path = {
    val dir = Paths.get(filename).getParent()
    if (dir != null)
      Files.createDirectories(dir)
    else
      null
  }

  def getListOfFiles(dir: String):List[File] = {
    val d = new File(dir)
    if (d.exists && d.isDirectory) {
      d.listFiles.filter(_.isFile).toList
    } else {
      List[File]()
    }
  }

  def splitListBy[T](pred:T=>Boolean, lst:List[T]):List[List[T]] = {
    lst match {
      case Nil => Nil
      case _ => {
        val (head, tail) = lst.span(pred)
        if (head.isEmpty)
          splitListBy[T](!pred(_), tail)
        else
          head::splitListBy[T](!pred(_), tail)
      }
    }
  }

  def pairingListBy[T](pred:T=>Boolean, lst:List[T]):List[(T,T)] = {
    lst match {
      case Nil => Nil
      case x if x.length < 2 => Nil
      case head::tail if (!pred(head)) => pairingListBy(pred, tail)
      case h1::h2::tail => (h1,h2)::pairingListBy(pred, tail)
    }
  }

  def blockingListBy[T](pred:T=>Boolean, lst:List[T]):List[List[T]] = {
    lst match {
      case Nil => Nil
      case _ => {
        val (head, tail) = lst.span(pred)
        head match {
          case Nil => blockingListBy[T](pred, tail.dropWhile(!pred(_)))
          case _ => head::blockingListBy[T](pred, tail.dropWhile(!pred(_)))
        }
      }
    }
  }

  def pairwiseBy[T](lst:List[T]):List[(T,T)] = {
    lst match {
      case Nil => Nil
      case x if x.length < 2 => Nil
      case head::tail => (head,tail.head)::pairwiseBy(tail)
    }
  }
}
