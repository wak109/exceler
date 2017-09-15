/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
package exceler.app

import org.eclipse.jetty.server.{Server, ServerConnector}
import org.eclipse.jetty.servlet.{DefaultServlet, ServletContextHandler}
import org.eclipse.jetty.webapp.WebAppContext
import org.scalatra.servlet.ScalatraListener

import org.scalatra.LifeCycle
import javax.servlet.ServletContext


object ServerLauncher {
  def run(port:Int) {

    val server = new Server()

    val connector = new ServerConnector(server)
    connector.setInheritChannel(true)
    connector.setPort(port)

    server.setConnectors(Array(connector))

    val context = new WebAppContext()
    context.setInitParameter(ScalatraListener.LifeCycleKey,
      "exceler.app.ScalatraBootstrap")
    context.setContextPath("/")
    context.setResourceBase("src/main/webapp")
    context.addEventListener(new ScalatraListener)
    context.addServlet(classOf[DefaultServlet], "/")

    server.setHandler(context)

    server.start
    server.join
  }
}
