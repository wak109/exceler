/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import org.scalatra.test.specs2._

import scala.util.{Try, Success, Failure}

class ExcelerServletTest extends MutableScalatraSpec {
  addServlet(classOf[ExcelerServlet], "/*")

  "GET / on ExcelerServlet" should {
    "return status 200" in {
      get("/") {
        status must_== 200
      }
    }
  }
}
