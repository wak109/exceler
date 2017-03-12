import org.scalatest.FunSuite

class ExcelerSuite extends FunSuite {
  
  test("Hello, world") {
    val cmdline = Exceler.parseCommandLine(Array("test.xlsx"))
    assert(1 === 1)
  }

}
