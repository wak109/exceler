/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */


import java.io._
import java.nio.file._

object CommonLib {

    object ImplicitConversions {

        implicit class ImplicitListConversion[T](val lst:List[T]) {
            def splitBy(pred:T=>Boolean) = splitListBy(pred, lst)
        }
    }

    def createParentDir(filename:String) : Path = {
        val dir = Paths.get(filename).getParent()
        if (dir != null)
            Files.createDirectories(dir)
        else
            null
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
}
