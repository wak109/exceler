/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.collection.JavaConverters._
import scala.util.Try

import java.io._
import java.net.URI
import java.nio.file._
import java.util.Properties


case class MyProperties(val properties:Properties) {

  def get(key:String):Option[String] = {
    Option(properties.getProperty(key))
      .map(Some(_))
      .getOrElse(
        Option(System.getProperties.getProperty(key)))
  }
}

object MyProperties {

  def getProperties(uri:URI):Option[Properties] = 
    (for {
      stream <- Try(uri.toURL.openStream())
    } yield {
      val properties = new Properties()
      properties.load(stream)
      stream.close()
      properties
    }).toOption

  def getProperties(key:String):Option[Properties] = 
    for {
      path <- Option(System.getProperties.getProperty(key))
      properties <- getProperties(new URI(path))
    } yield properties

  def apply(uri:URI):MyProperties = 
    new MyProperties(getProperties(uri).getOrElse(new Properties))

  def apply(key:String):MyProperties = 
    new MyProperties(getProperties(key).getOrElse(new Properties))

  def apply(key:String, uri:URI):MyProperties = 
    new MyProperties(getProperties(key)
      .getOrElse(getProperties(uri)
        .getOrElse(new Properties)))

  def apply(uri:URI, key:String):MyProperties = 
    new MyProperties(getProperties(uri)
      .getOrElse(getProperties(key)
        .getOrElse(new Properties)))
}
