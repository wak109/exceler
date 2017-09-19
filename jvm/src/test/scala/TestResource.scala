/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.test

trait TestResource {

  val testWorkbook1 = "test1.xlsx"

  def getURI(resource:String) =
    this.getClass.getClassLoader.getResource(resource).toURI
}
