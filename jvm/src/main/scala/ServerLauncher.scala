/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import org.eclipse.jetty.server.{Server,ServerConnector}
import org.eclipse.jetty.servlet.{DefaultServlet,ServletContextHandler}
import org.eclipse.jetty.webapp.WebAppContext

import org.scalatra.servlet.ScalatraListener
import org.scalatra.LifeCycle

import java.net.URI
import javax.servlet.ServletContext

object ServerLauncher {
  def run(port:Int) {

    val server = new Server()

    val connector = new ServerConnector(server)
    connector.setInheritChannel(true)
    connector.setPort(port)

    server.setConnectors(Array(connector))

    val context = new WebAppContext()
    context.setContextPath("/")
    val path = findWebResourceBase(this.getClass.getClassLoader)
    context.setResourceBase(path)
    context.addEventListener(new ScalatraListener)
    context.addServlet(classOf[DefaultServlet], "/")

    server.setHandler(context)

    server.start
    server.join
  }

  def findWebResourceBase(classLoader:ClassLoader) = {

    val webResourceRef = "WEB-INF/web.xml";
    val webXml = classLoader.getResource(webResourceRef)
    if (webXml != null)
      webXml.toURI().normalize().toString.replaceAll(webResourceRef + "$", "")
    else {
      "src/main/webapp"
    }
  }
}
