import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import java.io._

import org.apache.commons.cli.CommandLine
import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.HelpFormatter
import org.apache.commons.cli.{Option => CmdOption}
import org.apache.commons.cli.{Options => CmdOptions}
import org.apache.commons.cli.ParseException

import Exceler._

object Main {

    val description = """Scala CLI template"""
    val site = """https://github.com/wak109/scala_template"""

    def checkOptions(cl:CommandLine) : Unit = {
        if (cl.hasOption('h')) {
            printUsage()
        } else if (cl.getArgs().length == 0) {
            printUsage()
        } else {
            excel(cl.getArgs()(0))
        }
    }

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

    def parseCommandLine(args:Array[String]) : CommandLine = {
        val parser = new DefaultParser()

        return parser.parse(makeOptions(), args)
    }

    def makeOptions() : CmdOptions = {
        val options = new CmdOptions()
        options.addOption(CmdOption.builder("h").desc("Show help").build())

        return options
    }

    def main(args:Array[String]) : Unit  = {
        allCatch withTry { parseCommandLine(args) } match {
            case Success(cl) => checkOptions(cl)
            case Failure(e) => println(e.getMessage()); printUsage()
        }
    }
}
