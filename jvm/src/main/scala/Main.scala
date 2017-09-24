/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}
import scala.xml.Elem

import java.io._

import org.apache.commons.cli.CommandLine
import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.HelpFormatter
import org.apache.commons.cli.{Option => CmdOption}
import org.apache.commons.cli.{Options => CmdOptions}
import org.apache.commons.cli.ParseException

import exceler.excel._

sealed trait Config {
  val excelDir:String
}

object Config {
  def apply():Config = Main.getConfig
}


object DEFAULT {
  val EXCEL_DIR = "."
  val PORT = 8080
}

object Main {

  private var config:Config = null

  def getConfig() = config
  class ConfigImpl(val excelDir:String) extends Config


  val description = """Scala Excel"""
  val site = """https://github.com/wak109/exceler"""

  def stripClassName(clsname:String):String = {
    val Pattern = """^(.*)\$$""".r
    clsname match {
      case Pattern(m) => m
      case x => x
    }
  }

  def printUsage() : Unit = {
    val formatter = new HelpFormatter()

    formatter.printHelp(
      stripClassName(this.getClass.getCanonicalName),
      "\n" + this.description + "\n\noptions:\n",
      makeOptions(),
      "\nWeb: " + this.site,
      true)
  }

  def makeOptions() : CmdOptions = {
    val options = new CmdOptions()

    options.addOption("h", false, "Show help")
    options.addOption("p", true, "Listen port")
    options.addOption("d", true, "Excel directory")
    options.addOption("f", true, "Configuration file")

    return options
  }

  /**
   *
   */
  def parseCommandLine(args:Array[String])
        :Try[(List[String], Boolean, Int, String)] =
    Try {
      val parser = new DefaultParser()
      val cl = parser.parse(makeOptions(), args)
      (
        cl.getArgs.toList,
        cl.hasOption('h'),
        if (cl.hasOption('p')) cl.getOptionValue('p').toInt
          else DEFAULT.PORT,
        if (cl.hasOption('d')) cl.getOptionValue('d')
          else DEFAULT.EXCEL_DIR
      )
    }

  /**
   *
   */
  def main(args:Array[String]) : Unit  = 
    parseCommandLine(args) match {
      case Success(a) => a match {
        case (_,true,_,_) => printUsage() 
        case (Nil,_,port,dir) => {
          this.config = new ConfigImpl(dir)
          ServerLauncher.run(port)
        }
        case _ => printUsage()
      }
      case Failure(e) => println(e.getMessage); printUsage()
    }
}
