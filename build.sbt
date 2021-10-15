name := "ExcelSearch"
version := "0.0.1"

scalaVersion := "3.0.2"
scalacOptions ++= Seq("-deprecation")

crossPaths := false

scalacOptions ++= Seq("-encoding", "UTF-8")

autoScalaLibrary := false

libraryDependencies ++= Seq(
  "org.scala-lang" %% "scala3-library" % scalaVersion.value,
)

libraryDependencies ++= Seq(
//  "org.apache.poi" % "poi" % "5.0.0",
  "org.apache.poi" % "poi-ooxml" % "5.0.0"
    exclude("com.github.virtuald", "*")
    exclude("commons-codec", "*")
    exclude("de.rototor.pdfbox", "*")
    exclude("org.apache.xmlgraphics", "*")
    exclude("org.apache.santuario", "*")
    exclude("org.apache.pdfbox", "*")
    exclude("org.bouncycastle", "*")
    exclude("org.slf4j", "*")
    exclude("xml-apis", "*")
)

libraryDependencies ++= Seq(
  "junit" % "junit" % "4.13.2" % Test,
  "com.novocode" % "junit-interface" % "0.11" % Test,
)
