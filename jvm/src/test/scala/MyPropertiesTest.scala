/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import exceler.test._
import org.scalatest.FunSuite

class ExcelerConfigTest extends FunSuite with TestResource {
  
  test("MyProperties") {
    val prop = MyProperties("hehe", getURI(testProperties1))
    assert(prop.get("test.item").get == "HelloWorld")
  }
}
