import org.eclipse.jetty.server.{Server, ServerConnector}
import org.eclipse.jetty.servlet.{DefaultServlet, ServletContextHandler}
import org.eclipse.jetty.webapp.WebAppContext
import org.scalatra.servlet.ScalatraListener

import org.scalatra.LifeCycle
import javax.servlet.ServletContext


object JettyLauncher {
  def main(args: Array[String]) {
    val port = if(System.getenv("PORT") != null) System.getenv("PORT").toInt else 8080

    val server = new Server()

    val connector = new ServerConnector(server)
    connector.setInheritChannel(true)
    connector.setPort(port)

    server.setConnectors(Array(connector))

    val context = new WebAppContext()
    context.setContextPath("/")
    context.setResourceBase("src/main/webapp")
    context.addEventListener(new ScalatraListener)
    context.addServlet(classOf[DefaultServlet], "/")

    server.setHandler(context)

    server.start
    server.join
  }
}
