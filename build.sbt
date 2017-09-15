name := "exceler"
version := "0.3.0"
scalaVersion := "2.12.3"
scalacOptions ++= Seq("-unchecked", "-deprecation", "-feature")
libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "3.16",
  "org.apache.poi" % "poi-ooxml" % "3.16",
  "org.apache.poi" % "poi-ooxml-schemas" % "3.16",
  "org.scalactic" %% "scalactic" % "3.0.4",
  "org.scalatest" %% "scalatest" % "3.0.4" % "test",
  "commons-cli" % "commons-cli" % "1.4",
  "org.scalatra" %% "scalatra" % "2.5.1",
  "org.scalatra" %% "scalatra-specs2" % "2.5.1"  % "test",
  "ch.qos.logback" % "logback-classic" % "1.1.3" % "runtime",
  "org.eclipse.jetty" %  "jetty-webapp" % "9.4.7.RC0" % "compile;container",
  "javax.servlet" % "javax.servlet-api" % "3.1.0" % "provided",
  "junit" % "junit" % "4.12" % "test"
)

enablePlugins(JettyPlugin)

mainClass in (Compile, packageBin) := Some("exceler.ServerMain")
