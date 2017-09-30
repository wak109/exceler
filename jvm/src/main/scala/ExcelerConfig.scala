/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.collection.JavaConverters._
import scala.xml.Elem

import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.scalatra._

import java.io._
import java.net.URI
import java.nio.file._
import java.util.Properties

import javax.servlet.http.HttpServlet

import exceler.common._
import exceler.excel._
import excellib.ImplicitConversions._

import CommonLib._

object ExcelerConfig {

  val CONF_KEY = "exceler.conf"
  val DIR_KEY = "exceler.dir"
  val PORT_KEY = "exceler.port"

  val DEFAULT_CONF = "file:./exceler.conf"
  val DEFAULT_DIR ="."
  val DEFAULT_PORT ="9000"

  lazy val properties = MyProperties(CONF_KEY, new URI(DEFAULT_CONF))

  lazy val dir:String =
    properties.get(DIR_KEY).getOrElse(DEFAULT_DIR)

  lazy val port:Int =
    properties.get(PORT_KEY).getOrElse(DEFAULT_PORT).toInt
}
