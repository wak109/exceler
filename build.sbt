lazy val scalatraVersion = "2.5.1"

lazy val root = (project in file(".")).settings(
//  organization := "com.example",
  name := "exceler",
  version := "0.2.0",
  scalaVersion := "2.12.3",
  scalacOptions ++= Seq("-unchecked", "-deprecation", "-feature"),
  libraryDependencies ++= Seq(
    "org.apache.poi" % "poi" % "3.16",
    "org.apache.poi" % "poi-ooxml" % "3.16",
    "org.apache.poi" % "poi-ooxml-schemas" % "3.16",
    "org.scalactic" %% "scalactic" % "3.0.3",
    "org.scalatest" %% "scalatest" % "3.0.3" % "test",
    "commons-cli" % "commons-cli" % "1.4",
    "org.scalatra" %% "scalatra" % "2.5.1",
    "org.scalatra" %% "scalatra-specs2" % "2.5.1"  % "test",
    "ch.qos.logback" % "logback-classic" % "1.1.3" % "runtime",
    "org.eclipse.jetty" %  "jetty-webapp" % "9.4.6.v20170531" % "compile;container",
    "javax.servlet" % "javax.servlet-api" % "3.1.0" % "provided"
  )
).settings(jetty(): _*)

mainClass in (Compile, packageBin) := Some("Main")
