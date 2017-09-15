/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

import org.scalatest.FunSuite
import exceler.tablex._

class TableXTest extends FunSuite with TestResource {

  class CellTest(
    override val row:Int,
    override val col:Int,
    override val value:String
  ) extends CellX[String] {
  }

  class RectTest(
    override val row:Int,
    override val col:Int,
    override val height:Int,
    override val width:Int,
    override val value:String
  ) extends CellTest(row,col,value) with RangeX[String]

  test("TableX.apply") {
    val table = TableX(Seq(
      new CellTest(0, 0, "Cell00"),
      new CellTest(0, 1, "Cell01"),
      new CellTest(1, 0, "Cell10"),
      new CellTest(1, 1, "Cell11"),
      new CellTest(300, 400, "Cell34"),
      new CellTest(300, 5000, "Cell35")
    ))

    assert(table(0)(0).value == "Cell00")
    assert(table(2)(0).value == "Cell34")
  }

  test("RangeX") {
    val rect = new RectTest(0,0,3,4,"Rect")

    assert(rect.cellList.length == 12)
  }

  test("TableX.toCompact") {
    val rect = new RectTest(0,0,3,4,"Rect")

    val table = TableX.toCompact(TableX(rect.cellList))

    assert(table.length == 1)
    assert(table(0).length == 1)
    assert(table(0)(0).value == "Rect")
  }

  test("TableX.toArray") {
    val rect = new RectTest(0,0,3,4,"Rect")

    val table = TableX.toArray(
      TableX.toCompact(TableX(rect.cellList)))

    assert(table.length == 3)
    assert(table(0).length == 4)
    assert(table(2)(3).value == "Rect")
  }
}
