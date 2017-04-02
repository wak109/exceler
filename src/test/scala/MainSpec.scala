import org.scalatest.FunSuite

class MainSuite extends FunSuite {
  
  test("Hello, world") {
    val cmdline = Main.parseCommandLine(Array("test.xlsx"))
    assert(1 === 1)
  }

}
