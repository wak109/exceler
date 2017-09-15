/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

import org.scalatest.FunSuite
import exceler.tablex._

class AbcTableTest extends FunSuite with TestResource {

  class CellTest(
    override val row:Int,
    override val col:Int,
    override val value:String
  ) extends AbcCell[String] {
  }

  class RectTest(
    override val row:Int,
    override val col:Int,
    override val height:Int,
    override val width:Int,
    override val value:String
  ) extends CellTest(row,col,value) with AbcRange[String]

  test("AbcTable.apply") {
    val table = AbcTable(Seq(
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

  test("AbcRange") {
    val rect = new RectTest(0,0,3,4,"Rect")

    assert(rect.cellList.length == 12)
  }

  test("AbcTable.toCompact") {
    val rect = new RectTest(0,0,3,4,"Rect")

    val table = AbcTable.toCompact(AbcTable(rect.cellList))

    assert(table.length == 1)
    assert(table(0).length == 1)
    assert(table(0)(0).value == "Rect")
  }

  test("AbcTable.toArray") {
    val rect = new RectTest(0,0,3,4,"Rect")

    val table = AbcTable.toArray(
      AbcTable.toCompact(AbcTable(rect.cellList)))

    assert(table.length == 3)
    assert(table(0).length == 4)
    assert(table(2)(3).value == "Rect")
  }
}
