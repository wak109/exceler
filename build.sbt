// vim: set ts=2 et sw=2 sts=2 fileencoding=utf-8:

lazy val excelerName         = "exceler"
lazy val excelerOrganization = "exceler"
lazy val excelerVersion      = "0.6.0"

lazy val scalaLangVersion    = "2.12.4"
lazy val scalaJSDomVersion   = "0.9.2"
lazy val scalaJSJQueryVersion   = "0.9.2"
lazy val scalaJSReactVersion = "1.1.1"
lazy val scalatraVersion     = "2.5.3"
lazy val scalaTestVersion    = "3.0.4"
lazy val scalaCSSVersion     = "0.5.3"
lazy val reactJSVersion      = "15.6.1"
lazy val apachePoiVersion    = "3.17"


lazy val root:Project = project.in(file(".")).aggregate(jvm, js)

lazy val commonSettings = Seq(
  name := excelerName,
  organization := excelerOrganization,
  scalaVersion := scalaLangVersion,
  version      := excelerVersion,
  scalacOptions ++= Seq(
    "-deprecation",
    "-feature",
    "-unchecked",
    "-Xlint",
    "-Ywarn-dead-code",
    "-Ywarn-numeric-widen",
    "-Ywarn-unused",
    "-Ywarn-unused-import",
    "-Ywarn-value-discard"
    // -Xfatal-warnings: treat WARNING as an ERROR
    //, "-Xfatal-warnings"
  )
)

lazy val jvm:Project = project.in(file("jvm"))
  .settings(
    commonSettings,
    libraryDependencies ++= Seq(
      "org.apache.poi" % "poi" % apachePoiVersion,
      "org.apache.poi" % "poi-ooxml" % apachePoiVersion,
      "org.apache.poi" % "poi-ooxml-schemas" % apachePoiVersion,
      "org.scalatra" %%% "scalatra" % scalatraVersion,
      "org.scalatra" %%% "scalatra-specs2" % scalatraVersion  % "test",
      "commons-cli" % "commons-cli" % "1.4",
      "ch.qos.logback" % "logback-classic" % "1.1.3" % "runtime",
      "org.eclipse.jetty" %  "jetty-webapp" % "9.4.8.v20171121" % "compile;container",
      "javax.servlet" % "javax.servlet-api" % "3.1.0" % "provided",
      "junit" % "junit" % "4.12" % "test",
      "org.scalactic" %%% "scalactic" % scalaTestVersion,
      "org.scalatest" %%% "scalatest" % scalaTestVersion % "test"
    ),
    resolvers += Classpaths.typesafeReleases,
    mainClass in (Compile, packageBin) := Some("Main"),
    unmanagedSourceDirectories in Compile ++= Seq(
      baseDirectory.value / "../shared"
    ),
    unmanagedResourceDirectories in Compile ++= Seq(
      baseDirectory.value / "src/main/webapp"
    )
  )
  .enablePlugins(JettyPlugin)
  .enablePlugins(ScalatraPlugin)

lazy val js:Project = project.in(file("js"))
  .settings(
    commonSettings,
    scalaJSUseMainModuleInitializer := true,
    jsEnv := new org.scalajs.jsenv.jsdomnodejs.JSDOMNodeJSEnv(),
    testFrameworks += new TestFramework("utest.runner.Framework"),
    libraryDependencies ++= Seq(
      "org.scala-js" %%% "scalajs-dom" % scalaJSDomVersion,
      "be.doeraene" %%% "scalajs-jquery" % scalaJSJQueryVersion,
      "com.github.japgolly.scalajs-react" %%% "core" % scalaJSReactVersion,
      "com.github.japgolly.scalajs-react" %%% "extra" % scalaJSReactVersion,
      "com.github.japgolly.scalacss" %%% "core" % scalaCSSVersion,
      "com.github.japgolly.scalacss" %%% "ext-react" % scalaCSSVersion,
      "com.lihaoyi" %%% "utest" % "0.6.0" % "test"
    ),
    unmanagedSourceDirectories in Compile ++= Seq(
      baseDirectory.value / "../shared"
    ),
    skip in packageJSDependencies := false,
    jsDependencies ++= Seq(
      "org.webjars" % "jquery" % "2.1.4" / "2.1.4/jquery.js",
    
      "org.webjars.bower" % "react" % reactJSVersion
        /        "react-with-addons.js"
        minified "react-with-addons.min.js"
        commonJSName "React",
    
      "org.webjars.bower" % "react" % reactJSVersion
        /         "react-dom.js"
        minified  "react-dom.min.js"
        dependsOn "react-with-addons.js"
        commonJSName "ReactDOM",
    
      "org.webjars.bower" % "react" % reactJSVersion
        /         "react-dom-server.js"
        minified  "react-dom-server.min.js"
        dependsOn "react-dom.js"
        commonJSName "ReactDOMServer"
    ),
    assemblyMergeStrategy in assembly := {
      case "JS_DEPENDENCIES" => MergeStrategy.discard
      case x =>
        val oldStrategy = (assemblyMergeStrategy in assembly).value
        oldStrategy(x)
    },
    crossTarget in (Compile, fullOptJS) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, fastOptJS) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, packageJSDependencies) := file("jvm/src/main/webapp/js"),
    crossTarget in (Compile, packageMinifiedJSDependencies) := file("jvm/src/main/webapp/js"),
    artifactPath in (Compile, fastOptJS) := ((crossTarget in (Compile, fastOptJS)).value / ((moduleName in fastOptJS).value + "-opt.js"))
  )
  .enablePlugins(ScalaJSPlugin)
