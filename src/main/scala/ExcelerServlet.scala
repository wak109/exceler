import org.scalatra._

class ExcelerServlet extends ScalatraServlet {

  get("/") {
    <html>
      <body>
        <h1>Hello, world!</h1>
        <p>Scalatra</p>
      </body>
    </html>
  }
}
