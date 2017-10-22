/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.annotation.meta.field
import scala.scalajs.js
import scala.scalajs.js.annotation._
import scala.scalajs.js.Dynamic.{global,literal}
import scala.scalajs.concurrent.JSExecutionContext.Implicits.runNow

import org.scalajs.jquery._

import org.scalajs.dom
import dom.document
import dom.ext.Ajax

import scalajsreact.template

import scalajsreact.template.css.AppCSS
import scalajsreact.template.routes.AppRouter

import scalajsreact.template.css.AppCSS
import scalajsreact.template.routes.AppRouter

import org.scalajs.dom


object ExcelerJS {

  def appendPar(targetNode: dom.Node, text: String): Unit = {
    val parNode = document.createElement("p")
    val textNode = document.createTextNode(text)
    parNode.appendChild(textNode)
    targetNode.appendChild(parNode)
  }

  def main(args: Array[String]): Unit = {
    val json = js.Dynamic.literal(
      url = "http://127.0.0.1:8080/api",
      `type` = "GET"
    ).asInstanceOf[JQueryAjaxSettings]
    val promise = jQuery.ajax(json)
    promise.done((data:String, textStatus:String, xhr:JQueryXHR) => {
      val xmlDoc = jQuery.parseXML(xhr.responseText)
      val nameList = jQuery(xmlDoc).find("book").each(
        (elem:dom.Element)=> jQuery("#main").append(
          "<li><p>" + elem.getAttribute("name") + "</p></li>"))
      AppCSS.load
      AppRouter.router().renderIntoDOM(dom.document.body)
    })
  }
}
