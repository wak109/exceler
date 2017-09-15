/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.app

import org.scalatra._
import javax.servlet.ServletContext

//import exceler.app.ExcelerServlet

class ScalatraBootstrap extends LifeCycle {
  override def init(context: ServletContext) {
    context.mount(new ExcelerServlet, "/*")
  }
}
