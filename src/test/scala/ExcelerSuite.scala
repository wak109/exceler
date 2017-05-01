import org.scalatest.FunSuite

class ExcelerSuite extends FunSuite {
  
  test("Create Excel Workbook") {
      val workbook = ExcelerWorkbook.create()
      ExcelerWorkbook.createSheet(workbook, "test")
      ExcelerWorkbook.saveAs(workbook, "hehe.xlsx")
  }
}
