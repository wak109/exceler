/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

import scala.util.{Try, Success, Failure}
import org.scalatest.FunSuite

import exceler._

class ServerMainSuite extends FunSuite {
  
  test("Main.parseCommandLine") {
    ServerMain.parseCommandLine(Array()) match {
      case Success((Nil, false, 8080, ".")) => assert(true)
      case _ => assert(false)
    }
    ServerMain.parseCommandLine(Array("-p", "9999")) match {
      case Success((Nil, false, 9999, ".")) => assert(true)
      case _ => assert(false)
    }
    ServerMain.parseCommandLine(Array("-p", "abcd")) match {
      case Failure(_) => assert(true)
      case _ => assert(false)
    }
  }
}
