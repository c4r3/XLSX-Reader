name := "XLSX-Reader"

version := "1.0.0"

scalaVersion := "2.12.12"

libraryDependencies ++= Seq(
  "org.scala-lang.modules" %% "scala-xml" % "1.3.0",
  "org.scalatest" %% "scalatest" % "3.2.0" % Test,
  "org.slf4j" % "slf4j-api" % "1.7.30",
  "org.slf4j" % "slf4j-log4j12" % "1.7.30"
)