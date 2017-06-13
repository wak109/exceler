/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */

import java.io._
import java.nio.file._

trait CommonLib {

    def createParentDir(filename:String) : Path = {
        val dir = Paths.get(filename).getParent()
        if (dir != null)
            Files.createDirectories(dir)
        else
            null
    }
}

object CommonLib extends CommonLib

