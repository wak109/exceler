import org.scalatest.FunSuite

class MainSuite extends FunSuite {
  
  test("Main.parseCommandLine") {
    val cmdline = Main.parseCommandLine(Array("test.xlsx"))
    assert(1 === 1)
  }

}
