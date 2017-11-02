/* vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8: */


val scalaJSReactVersion = "1.1.0"
val scalaCssVersion = "0.5.3"
val reactJSVersion = "15.6.1"

lazy val root:Project = project.in(file("."))
  .aggregate(jvm, js)

lazy val commonSettings = Seq(
    organization := "exceler",
    scalaVersion := "2.12.3",
    version      := "0.5.0",
    scalacOptions ++= Seq("-unchecked", "-deprecation", "-feature")
  )

lazy val jvm:Project = project.in(file("jvm"))
  .settings(
    commonSettings,
    name := "exceler",
    libraryDependencies ++= Seq(
      "org.apache.poi" % "poi" % "3.17",
      "org.apache.poi" % "poi-ooxml" % "3.17",
      "org.apache.poi" % "poi-ooxml-schemas" % "3.17",
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
    resolvers += Classpaths.typesafeReleases,
    mainClass in (Compile, packageBin) := Some("Main"),
    unmanagedSourceDirectories in Compile ++= Seq(
      baseDirectory.value / "../shared"
    ),
    unmanagedResourceDirectories in Compile ++= Seq(
      baseDirectory.value / "src/main/webapp"
    )/*,
    resourceGenerators in Compile += Def.task {
      val opt = baseDirectory.value /
        "src/main/webapp/js/exceler-fastopt.js"
      val map = baseDirectory.value /
        "src/main/webapp/js/exceler-fastopt.js.map"
      IO.copyFile((fastOptJS in Compile in js).value.data, opt)
      Seq(opt,map)
    }.taskValue,
    resources in Compile += (fastOptJS in Compile in js).value.data,
    cleanFiles ++= Seq(
      baseDirectory.value / "src/main/webapp/js/exceler-fastopt.js"
    )
*/
  )
  .enablePlugins(JettyPlugin)
  .enablePlugins(ScalatraPlugin)

lazy val js:Project = project.in(file("js"))
  .settings(
    commonSettings,
    name := "exceler",
    scalaJSUseMainModuleInitializer := true,
    //mainClass := Some(ExcelerJS),
    libraryDependencies ++= Seq(
      "org.scala-js" %%% "scalajs-dom" % "0.9.2",
      "be.doeraene" %%% "scalajs-jquery" % "0.9.2",
      "org.scalatest" %%% "scalatest" % "3.0.4" % "test",
      "be.doeraene" %%% "scalajs-jquery" % "0.9.2",
      "com.github.japgolly.scalajs-react" %%% "core" % scalaJSReactVersion,
      "com.github.japgolly.scalajs-react" %%% "extra" % scalaJSReactVersion,
      "com.github.japgolly.scalacss" %%% "core" % scalaCssVersion,
      "com.github.japgolly.scalacss" %%% "ext-react" % scalaCssVersion
    ),
    unmanagedSourceDirectories in Compile ++= Seq(
      baseDirectory.value / "../shared"
    ),
    jsDependencies ++= Seq(
      "org.webjars.npm" % "react" % reactJSVersion / "react-with-addons.js" commonJSName "React" minified "react-with-addons.min.js",
      "org.webjars.npm" % "react-dom" % reactJSVersion / "react-dom.js" commonJSName "ReactDOM" minified "react-dom.min.js" dependsOn "react-with-addons.js"
    ),
    skip in packageJSDependencies := false,
    crossTarget in (Compile, fullOptJS) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, fastOptJS) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, packageJSDependencies) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, packageMinifiedJSDependencies) := file("jvm/src/main/webapp/js"),
    artifactPath in (Compile, fastOptJS) := ((crossTarget in (Compile, fastOptJS)).value / ((moduleName in fastOptJS).value + "-opt.js"))
  )
  .enablePlugins(ScalaJSPlugin)
