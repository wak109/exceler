/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */

lazy val root = project.in(file("."))
  .aggregate(excelerJS, excelerJVM)
  .settings(
    publish := {},
    publishLocal := {},
    publishM2 := {}
  )

lazy val exceler = crossProject.in(file("."))
  .settings(
    inThisBuild(List(
      organization := "exceler",
      scalaVersion := "2.12.3",
      version      := "0.4.0"
    )),
    name := "exceler",
    libraryDependencies ++= Seq(
    )
  )
  .jvmSettings(
    scalacOptions ++= Seq("-unchecked", "-deprecation", "-feature"),
    libraryDependencies ++= Seq(
      "org.apache.poi" % "poi" % "3.16",
      "org.apache.poi" % "poi-ooxml" % "3.16",
      "org.apache.poi" % "poi-ooxml-schemas" % "3.16",
      "commons-cli" % "commons-cli" % "1.4",
      "org.scalatra" %%% "scalatra" % "2.5.1",
      "org.scalatra" %%% "scalatra-specs2" % "2.5.1"  % "test",
      "ch.qos.logback" % "logback-classic" % "1.1.3" % "runtime",
      "org.eclipse.jetty" %  "jetty-webapp" % "9.4.7.RC0" % "compile;container",
      "javax.servlet" % "javax.servlet-api" % "3.1.0" % "provided",
      "junit" % "junit" % "4.12" % "test",
      "org.scalactic" %%% "scalactic" % "3.0.4",
      "org.scalatest" %%% "scalatest" % "3.0.4" % "test"
    ),
    mainClass in (Compile, packageBin) := Some("exceler.app.Main")
  )
  .jsSettings(
    scalaJSUseMainModuleInitializer := true,
    libraryDependencies ++= Seq(
      "org.scala-js" %%% "scalajs-dom" % "0.9.2",
      "be.doeraene" %%% "scalajs-jquery" % "0.9.2",
      "org.scalatest" %%% "scalatest" % "3.0.4" % "test"
    )
  )
  .enablePlugins(JettyPlugin)

lazy val excelerJS = exceler.js
  .enablePlugins(ScalaJSPlugin)
lazy val excelerJVM = exceler.jvm
  .settings(
    (resources in Compile) += (fastOptJS in (excelerJS, Compile)).value.data
  )
