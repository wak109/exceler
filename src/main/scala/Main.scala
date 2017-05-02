/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import scala.util.control.Exception._
import scala.util.{Try, Success, Failure}

import java.io._

import org.apache.commons.cli.CommandLine
import org.apache.commons.cli.DefaultParser
import org.apache.commons.cli.HelpFormatter
import org.apache.commons.cli.{Option => CmdOption}
import org.apache.commons.cli.{Options => CmdOptions}
import org.apache.commons.cli.ParseException


object Main {

    val description = """Scala CLI template"""
    val site = """https://github.com/wak109/scala_template"""

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
        options.addOption(CmdOption.builder("h").desc("Show help").build())

        return options
    }

    def checkOptions(cl:CommandLine) : Either[String,CommandLine] = {
        if (cl.hasOption('h')) {
            return Left("help option")
        } else if (cl.getArgs.isEmpty) {
            return Left("No Command")
        } else {
            return Right(cl)
        }
    }

    def parseCommandLine(args:Array[String]) : Try[CommandLine] = {
        Try {
            val parser = new DefaultParser()
            parser.parse(makeOptions(), args)
        }
    }

    def main(args:Array[String]) : Unit  = {
        (
            for {
                cl <- parseCommandLine(args)
                cl <- checkOptions(cl).left.map(new Exception(_)).toTry
            } yield cl
        ) match {
            case Success(cl) => {
                cl.getArgs()(0) match {
                    case _ => printUsage()
                }
            }
            case Failure(e) => println(e.getMessage()); printUsage()
        }
    }
}
