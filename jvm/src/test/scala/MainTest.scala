/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.{Try, Success, Failure}
import org.scalatest.FunSuite

class MainSuite extends FunSuite {
  
  test("Main.parseCommandLine") {
    Main.parseCommandLine(Array()) match {
      case Success((Nil, false, 8080, ".")) => assert(true)
      case _ => assert(false)
    }
    Main.parseCommandLine(Array("-p", "9999")) match {
      case Success((Nil, false, 9999, ".")) => assert(true)
      case _ => assert(false)
    }
    Main.parseCommandLine(Array("-p", "abcd")) match {
      case Failure(_) => assert(true)
      case _ => assert(false)
    }
  }
}
