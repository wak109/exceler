package my.react.basic

/**
  * Sample code for scala.js react
  *
  * - VDOM
  * - Router
  *
  * Scala.js React: https://github.com/japgolly/scalajs-react
  */

import scala.concurrent.Future
import scala.concurrent.ExecutionContext.Implicits.global
import org.scalajs.dom
import dom.document
import dom.ext.Ajax
import japgolly.scalajs.react._
import japgolly.scalajs.react.vdom.html_<^._
import japgolly.scalajs.react.extra.router._
import org.scalajs.dom.raw.Document

import scala.util.{Failure, Success}


object ExcelerAPI {

  def convXmlToWorkbooks(xmlDoc:Document):List[String] = {
    val bookNodes = xmlDoc.getElementsByTagName("book")
    (0 until bookNodes.length).map(bookNodes.item(_)
      .attributes.getNamedItem("name").value)(collection.breakOut)
  }

  def fetchWorkbooks:Future[List[String]] =
    Ajax.get("api").map(xhr => convXmlToWorkbooks(xhr.responseXML))
}


object StateComponent {

  /**
    * State for MyItemList
    *
    * @param items List of string of item
    * @param newItem new item under editing
    */
  case class State
  (
    items:Vector[String],
    newItem:String
  )

  private class Backend($: BackendScope[Unit, State]) {

    def updateNewItem(e:ReactEventFromInput):Callback = {
      // NOTE:
      //
      // BAD CODE:
      // $.modState(_.copy(newItem = e.target.value))
      //
      // REASON:
      // $.modState(...), which is based on the JS react "setState", works asynchronously.
      // "e.target.value", also which is base on JS.
      // So, when the JS "setState" is actually executed asynchronously,
      // the call for "e.target.value" doesn't work as expected.
      //
      val value = e.target.value
      $.modState(_.copy(newItem = value))
    }

    def addNewItem:Callback =
      $.modState((s: State) => s.copy(items = s.items :+ s.newItem))

    def start:Callback = Callback.future(ExcelerAPI.fetchWorkbooks.map(
      (workbooks:List[String]) => $.modState(
        (s:State)=>s.copy(items = s.items ++ workbooks))
    ))

    def render(s:State): VdomElement =
      <.div(
        MyItemList(s.items, updateNewItem, addNewItem)
      )
  }

  private val component =
    ScalaComponent.builder[Unit]("StateComponent")
      .initialState(State(Vector(), ""))
      .renderBackend[Backend]
      .componentDidMount(_.backend.start)
      .build

  def apply():VdomElement = component()
}

/**
  * Factory for MyItemList component
  */
object MyItemList {

  /**
    * State for MyItemList
    *
    * @param items List of string of item
    */
  case class Props
  (
    items:Vector[String],
    updateNewItem:ReactEventFromInput=>Callback,
    addNewItem:Callback
  )

  private class Backend($: BackendScope[Props,Unit]) {

    def render(p:Props): VdomElement =
      <.div(
        MyInput(p.updateNewItem, p.addNewItem),
        <.div(p.items.length, " items found:"),
        <.ol(p.items.toTagMod(i => MyItem(i)))
      )
  }

  private val component =
    ScalaComponent.builder[Props]("MyItemList")
      .renderBackend[Backend]
      .build

  def apply
  (
    items:Vector[String],
    updateNewItem:ReactEventFromInput=>Callback,
    addNewItem:Callback
  ):VdomElement = component(Props(items,updateNewItem,addNewItem))
}


object MyItem {

  private case class Props(item:String)

  private class Backend($: BackendScope[Props,Unit]) {
    def render(p: Props): VdomElement =
      <.li(p.item)
  }

  private val component =
    ScalaComponent.builder[Props]("MyItem")
      .stateless
      .renderBackend[Backend]
      .build

  def apply(item:String):VdomElement = component(Props(item))
}

/**
  * Factory for MyInput component
  */
object MyInput {

  /**
    * Props for MyInput component
    *
    * @param onTextUpdate callback whenever input text is updated
    * @param onButtonClick callback when clicking 'Add' button
    */
  private case class Props
  (
    onTextUpdate:ReactEventFromInput=>Callback,
    onButtonClick:Callback
  )

  /**
    * Backend for MyInput component
    *
    * @param $ Props for MyInput component
    */
  private class Backend($:BackendScope[Props,Unit]) {

    /**
      * Render MyInput component
      *
      * This method is called whenever Props is updated.
      * No need to call
      *
      * @param p Props for MyInput
      * @return VDOM for MyInput
      */
    def render(p:Props) =
      <.div(
        <.input(
          ^.`type`    := "text",
          ^.onChange ==> p.onTextUpdate
        ),
        <.button(
          ^.onClick --> p.onButtonClick,
          "Add"
        )
      )
  }

  private val component =
    ScalaComponent.builder[Props]("MyInput")
      .stateless
      .renderBackend[Backend]
      .build

  def apply(
            onTextUpdate:ReactEventFromInput=>Callback,
            onButtonClick:Callback
           ):VdomElement =
    component(Props(onTextUpdate,onButtonClick))
}

/**
  * Factory for router component
  */
object MyRouter {

  sealed trait MyPages
  case object Home extends MyPages
  case object About extends MyPages

  /**
    * Get the head part of URL before "#"
    *
    * @param url URL
    * @return BaseUrl
    */
  private def getBaseUrl(url:String):BaseUrl = BaseUrl(url.takeWhile(_ != '#'))

  /**
    * The router configuration, based on 2 rules below
    *
    *   - none, notFound -> render MyItemList
    *   - '#about' -> render <h1>Router test</h1>
    */
  private val routerConfig = RouterConfigDsl[MyPages].buildConfig { dsl =>
    import dsl._

    (emptyRule
      | staticRoute(root, Home) ~> render( StateComponent() )
      | staticRoute("#about", About) ~> render( <.h1("Router test"))
      )
      .notFound(redirectToPage(Home)(Redirect.Replace))
  }

  /**
    * Create a router component
    *
    * @param url URL of this application
    * @return router component based for the given URL
    */
  def apply(url:String) = Router(getBaseUrl(url), routerConfig)()
}

object MyReactBasicMain {

  def setup(node:dom.Element, url:String):Unit = {
    /**
      * Render the VDOM component directly if you don't need Router
      *
      * Example: MyItemList().renderIntoDOM(node)
      */
    MyRouter(url).renderIntoDOM(node)
  }

  def main(args: Array[String]): Unit = {
    setup(
      document.getElementById("exceler-app"),
      dom.window.location.href
    )
  }
}
