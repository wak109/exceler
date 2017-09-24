scalacOptions ++= Seq("-unchecked", "-deprecation", "-feature", "-language:postfixOps")

addSbtPlugin("org.scala-js" % "sbt-scalajs" % "0.6.20")
addSbtPlugin("org.scalatra.sbt" % "sbt-scalatra" % "1.0.0")
addSbtPlugin("com.eed3si9n" % "sbt-assembly" % "0.14.5")
