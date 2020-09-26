name := "ExcelSearch"
version := "0.0.1"

scalaVersion := "2.13.3"
scalacOptions ++= Seq("-deprecation")

crossPaths := false

scalacOptions ++= Seq("-encoding", "UTF-8")

autoScalaLibrary := false

libraryDependencies ++= Seq(
  "org.scala-lang" % "scala-library" % scalaVersion.value,
)

libraryDependencies ++= Seq(
  "org.apache.poi" % "poi" % "4.1.2",
  "org.apache.poi" % "poi-ooxml" % "4.1.2",
)

libraryDependencies ++= Seq(
  "junit" % "junit" % "4.13" % Test,
  "com.novocode" % "junit-interface" % "0.11" % Test,
)
