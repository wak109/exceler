package exceler.js

import scala.scalajs.js
import scala.scalajs.js.Dynamic.{global,literal}
import scala.scalajs.js.JSApp
import scala.scalajs.concurrent.JSExecutionContext.Implicits.runNow

import org.scalajs.dom
import dom.document
import dom.ext.Ajax

object ExcelerJS extends JSApp {
  def appendPar(targetNode: dom.Node, text: String): Unit = {
    val parNode = document.createElement("p")
    val textNode = document.createTextNode(text)
    parNode.appendChild(textNode)
    targetNode.appendChild(parNode)
  }
  def main(): Unit = {
    appendPar(document.body, "Hello, world!!!!!")
    Ajax.get("http://127.0.0.1:8080/api").foreach {
      xhr => println(xhr.responseText)
    }
  }
}
