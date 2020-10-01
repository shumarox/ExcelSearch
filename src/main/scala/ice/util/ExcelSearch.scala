package ice.util

import java.io.{BufferedOutputStream, File, FileOutputStream, IOException}
import java.nio.file._
import java.nio.file.attribute.BasicFileAttributes
import java.text.SimpleDateFormat

import ice.util.ExcelSearch.r1c1ToA1
import org.apache.poi.common.usermodel.HyperlinkType
import org.apache.poi.hssf.usermodel.{HSSFClientAnchor, HSSFShapeGroup, HSSFSimpleShape}
import org.apache.poi.ss.formula.eval.ErrorEval
import org.apache.poi.ss.usermodel._
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.{XSSFClientAnchor, XSSFShapeGroup, XSSFSimpleShape}

import scala.collection.mutable
import scala.util.matching.Regex
import scala.util.{Failure, Success, Try, Using}

sealed case class MatchedType(order: Int)

object MatchedType {

  object BookName extends MatchedType(0)

  object SheetName extends MatchedType(1)

  object Cell extends MatchedType(2)

  object Comment extends MatchedType(3)

  object Shape extends MatchedType(4)

  object Error extends MatchedType(5)

}

abstract class MatchedInfo(val matchedType: MatchedType, val file: File, val sheetName: String, val name: String, val rowIndex: Int, val columnIndex: Int, val text: String) {
  def location: String = r1c1ToA1(rowIndex, columnIndex)

  def url: String = {
    val fileNameString = file.getAbsolutePath.replaceAll("\\\\", "/").replaceAll(" ", "%20")

    val locationString =
      if (sheetName == null || sheetName.isEmpty) {
        ""
      } else {
        s"#'${sheetName.replaceAll(" ", "%20")}'!$location"
      }

    s"file:///$fileNameString$locationString"
  }

  override def toString: String = s"${file.getParent},${file.getName},$sheetName,$name,$location,$text"
}

class MatchedBookNameInfo(file: File, text: String)
  extends MatchedInfo(MatchedType.BookName, file, "", "ブック名", 0, 0, text)

class MatchedSheetNameInfo(file: File, sheetName: String, text: String)
  extends MatchedInfo(MatchedType.SheetName, file, sheetName, "シート名", 0, 0, text)

class MatchedCellInfo(file: File, sheetName: String, rowIndex: Int, columnIndex: Int, text: String)
  extends MatchedInfo(MatchedType.Cell, file, sheetName, r1c1ToA1(rowIndex, columnIndex), rowIndex, columnIndex, text)

class MatchedCommentInfo(file: File, sheetName: String, rowIndex: Int, columnIndex: Int, text: String)
  extends MatchedInfo(MatchedType.Comment, file, sheetName, "コメント " + r1c1ToA1(rowIndex, columnIndex), rowIndex, columnIndex, text)

class MatchedShapeInfo(file: File, sheetName: String, val shapeName: String, rowIndex: Int, columnIndex: Int, text: String)
  extends MatchedInfo(MatchedType.Shape, file, sheetName, shapeName, rowIndex, columnIndex, text)

class ErrorInfo(file: File, text: String)
  extends MatchedInfo(MatchedType.Error, file, "", "エラー", 0, 0, text)

object ExcelSearch {
  def main(args: Array[String]): Unit = {
    if (args.length == 0 || args.length > 3 || args(0) == null || args(0).trim().isEmpty || args(1) == null || args(1).isEmpty) {
      System.err.println("Requires target directory and regular expression arguments")
      System.exit(1)
    }

    val result = new ExcelSearch().search(Paths.get(args(0)), new Regex(args(1)))

    if (args.length == 3) {
      val resultFile = args(2)

      if (resultFile != null && resultFile.nonEmpty) {
        createResultBook(new File(resultFile), result)

        if (File.separatorChar == '\\') {
          Runtime.getRuntime.exec(s"cmd /c start $resultFile")
        }
      }
    }
  }

  def printMatchedInfo(matchedInfo: MatchedInfo): Unit = {
    System.out.println(
      matchedInfo.toString
        .replaceAll("\r", "\\\\r")
        .replaceAll("\n", "\\\\n")
    )
  }

  def createResultBook(file: File, result: Array[MatchedInfo]): Unit = {
    val workbook = WorkbookFactory.create(true)
    val creationHelper = workbook.getCreationHelper

    val sheet = workbook.createSheet

    var rowIndex = 0

    val font = workbook.createFont
    font.setColor(IndexedColors.BLUE.getIndex)
    font.setUnderline(Font.U_SINGLE)
    val style = workbook.createCellStyle
    style.setFont(font)

    result.foreach { matchedInfo =>
      val row = sheet.createRow(rowIndex)

      var cell: Cell = null

      cell = row.createCell(0)
      cell.setCellValue(matchedInfo.file.getParent)

      cell = row.createCell(1)
      cell.setCellValue(matchedInfo.file.getName)

      cell = row.createCell(2)
      cell.setCellValue(matchedInfo.sheetName)

      cell = row.createCell(3)
      val link = creationHelper.createHyperlink(HyperlinkType.URL)
      link.setAddress(matchedInfo.url)
      cell.setHyperlink(link)
      cell.setCellValue(matchedInfo.name)
      cell.setCellStyle(style)

      cell = row.createCell(4)
      cell.setCellValue(matchedInfo.text)

      rowIndex += 1
    }

    Using.resource(new BufferedOutputStream(new FileOutputStream(file))) { outputStream =>
      workbook.write(outputStream)
    }
  }

  def r1c1ToA1(rowIndex: Int, columnIndex: Int): String = new CellReference(rowIndex, columnIndex).formatAsString()

  def getCellValue(cell: Cell): String = {
    import CellType._

    val cellType =
      cell.getCellType match {
        case FORMULA => cell.getCachedFormulaResultType
        case t => t
      }

    cellType match {
      case BLANK =>
        ""
      case STRING =>
        cell.getStringCellValue
      case NUMERIC =>
        getNumericCellValue(cell)
      case BOOLEAN =>
        String.valueOf(cell.getBooleanCellValue).toUpperCase
      case ERROR =>
        ErrorEval.getText(cell.getErrorCellValue)
      case _ =>
        cell.getStringCellValue
    }
  }

  def getNumericCellValue(cell: Cell): String = {
    if (DateUtil.isCellDateFormatted(cell)) {
      val format = cell.getCellStyle.getDataFormatString

      if (format.contains("h")) {
        if (format.contains("y")) {
          new SimpleDateFormat("yyyy/MM/dd HH:mm:ss").format(cell.getDateCellValue)
        } else {
          new SimpleDateFormat("HH:mm:ss").format(cell.getDateCellValue)
        }
      } else {
        new SimpleDateFormat("yyyy/MM/dd").format(cell.getDateCellValue)
      }
    } else {
      String.valueOf(cell.getNumericCellValue).replaceAll("\\.0$", "")
    }
  }
}

class ExcelSearch {

  import ExcelSearch._

  private val resultBuffer = mutable.ArrayBuffer[MatchedInfo]()

  private def addMatchedInfo(matchedInfo: MatchedInfo): Unit = {
    printMatchedInfo(matchedInfo)
    resultBuffer += matchedInfo
  }

  def search(path: Path, regex: Regex): Array[MatchedInfo] = {
    val fileSearcher: FileVisitor[Path] = new FileVisitor[Path] {
      override def postVisitDirectory(dir: Path, exc: IOException) = FileVisitResult.CONTINUE

      override def preVisitDirectory(dir: Path, attrs: BasicFileAttributes) = FileVisitResult.CONTINUE

      override def visitFile(file: Path, attrs: BasicFileAttributes): FileVisitResult = {
        if (file.toFile.getName.matches(".*\\.xls[xm]?$")) search(file.toFile, regex)
        FileVisitResult.CONTINUE
      }

      override def visitFileFailed(file: Path, exc: IOException) = FileVisitResult.CONTINUE
    }

    Files.walkFileTree(path, fileSearcher)

    resultBuffer.toArray
  }

  private def search(file: File, regex: Regex): Unit = {
    Using(WorkbookFactory.create(file, null, true)) { workbook =>
      val bookName = file.getName

      if (regex.findFirstIn(bookName).nonEmpty) {
        addMatchedInfo(new MatchedBookNameInfo(file, bookName))
      }

      workbook.forEach(sheet => {
        val sheetName = sheet.getSheetName

        if (regex.findFirstIn(sheetName).nonEmpty) {
          addMatchedInfo(new MatchedSheetNameInfo(file, sheetName, sheetName))
        }

        sheet.forEach(row => row.forEach(cell => {
          Try(getCellValue(cell)).foreach { value =>
            if (value.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(value).nonEmpty) {
              addMatchedInfo(new MatchedCellInfo(file, sheetName, cell.getRowIndex, cell.getColumnIndex, value))
            }
          }

          Try(cell.getCellComment.getString.getString).foreach { comment =>
            if (comment.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(comment).nonEmpty) {
              addMatchedInfo(new MatchedCommentInfo(file, sheetName, cell.getRowIndex, cell.getColumnIndex, comment))
            }
          }
        }))

        val drawingPatriarch = sheet.getDrawingPatriarch

        if (drawingPatriarch != null) {
          drawingPatriarch.forEach(shape => {
            def processShape(shapeName: String, rowIndex: Int, columnIndex: Int, value: String): Unit = {
              if (value.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(value).nonEmpty) {
                addMatchedInfo(new MatchedShapeInfo(file, sheetName, shapeName, rowIndex, columnIndex, value))
              }
            }

            walkShape(shape, 0, 0, processShape)
          })
        }
      })
    } match {
      case Success(_) =>
      case Failure(ex) =>
        System.err.println("ERROR: " + file.getAbsolutePath)
        ex.printStackTrace()
        addMatchedInfo(new ErrorInfo(file, ex.toString))
    }
  }

  private def walkShape(shape: Any, ancestorRowIndex: Int, ancestorColumnIndex: Int, processShape: (String, Int, Int, String) => ()): Unit = {
    def getRowColumnIndex(shape: Shape): (Int, Int) =
      shape.getAnchor match {
        case null => (ancestorRowIndex, ancestorColumnIndex)
        case anchor: XSSFClientAnchor => (anchor.getRow1, anchor.getCol1.toInt)
        case anchor: HSSFClientAnchor => (anchor.getRow1, anchor.getCol1.toInt)
        case _ => (ancestorRowIndex, ancestorColumnIndex)
      }

    shape match {
      case shape: XSSFSimpleShape =>
        val (rowIndex, columnIndex) = getRowColumnIndex(shape)
        processShape(shape.getShapeName, rowIndex, columnIndex, shape.getText)
      case shape: HSSFSimpleShape =>
        val shapeName = shape.getShapeName
        if (shapeName != null) {
          val (rowIndex, columnIndex) = getRowColumnIndex(shape)
          val text = Try(shape.getString.getString).getOrElse("")
          processShape(shapeName, rowIndex, columnIndex, text)
        }
      case group: XSSFShapeGroup =>
        val (rowIndex, columnIndex) = getRowColumnIndex(group)
        group.forEach(walkShape(_, rowIndex, columnIndex, processShape))
      case group: HSSFShapeGroup =>
        val (rowIndex, columnIndex) = getRowColumnIndex(group)
        group.forEach(walkShape(_, rowIndex, columnIndex, processShape))
      case _ =>
    }
  }

}
