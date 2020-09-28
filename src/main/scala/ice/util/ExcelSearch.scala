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
import scala.util.{Try, Using}

sealed case class MatchedType(order: Int)

object MatchedType {

  object BookName extends MatchedType(0)

  object SheetName extends MatchedType(1)

  object Cell extends MatchedType(2)

  object Comment extends MatchedType(3)

  object Shape extends MatchedType(4)

}

abstract class MatchedInfo(val matchedType: MatchedType, val file: File, val sheetName: String, val name: String, val row: Int, val col: Int, val text: String) {
  def location: String = r1c1ToA1(row, col)

  override def toString: String = s"${file.getAbsolutePath},$sheetName,$name,$location,$text"
}

class MatchedBookNameInfo(file: File, text: String)
  extends MatchedInfo(MatchedType.BookName, file, "", "ブック名", 0, 0, text)

class MatchedSheetNameInfo(file: File, sheetName: String, text: String)
  extends MatchedInfo(MatchedType.SheetName, file, sheetName, "シート名", 0, 0, text)

class MatchedCellInfo(file: File, sheetName: String, row: Int, col: Int, text: String)
  extends MatchedInfo(MatchedType.Cell, file, sheetName, r1c1ToA1(row, col), row, col, text)

class MatchedCommentInfo(file: File, sheetName: String, row: Int, col: Int, text: String)
  extends MatchedInfo(MatchedType.Comment, file, sheetName, "コメント " + r1c1ToA1(row, col), row, col, text)

class MatchedShapeInfo(file: File, sheetName: String, val shapeName: String, row: Int, col: Int, text: String)
  extends MatchedInfo(MatchedType.Shape, file, sheetName, shapeName, row, col, text)

object ExcelSearch {
  def main(args: Array[String]): Unit = {
    if (args.length == 0 || args.length > 3 || args(0) == null || args(0).trim().isEmpty || args(1) == null || args(1).isEmpty) {
      System.err.println("Requires target directory and regular expression arguments")
      System.exit(1)
    }

    val result = new ExcelSearch().search(args(0), new Regex(args(1)))

    if (args.length != 3 || args(2) == null || args(2).trim().isEmpty) {
      printResult(result)
    } else {
      createResultBook(new File(args(2)), result)

      if (File.separatorChar == '\\') {
        Runtime.getRuntime.exec(s"cmd /c start ${args(2)}")
      }
    }
  }

  private def printResult(result: Array[MatchedInfo]): Unit = {
    result.foreach { matchInfo =>
      System.out.println(
        matchInfo.toString
          .replaceAll("\r", "\\\\r")
          .replaceAll("\n", "\\\\n")
      )
    }
  }

  private def createResultBook(file: File, result: Array[MatchedInfo]): Unit = {
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
      cell.setCellValue(matchedInfo.file.getAbsolutePath)

      cell = row.createCell(1)
      cell.setCellValue(matchedInfo.sheetName)

      cell = row.createCell(2)
      val link = creationHelper.createHyperlink(HyperlinkType.URL)

      val fileNameString = matchedInfo.file.getAbsolutePath.replaceAll("\\\\", "/").replaceAll(" ", "%20")

      val locationString =
        if (matchedInfo.sheetName == null || matchedInfo.sheetName.isEmpty) {
          ""
        } else {
          s"#'${matchedInfo.sheetName.replaceAll(" ", "%20")}'!${matchedInfo.location}"
        }

      link.setAddress(s"file:///$fileNameString$locationString")
      cell.setHyperlink(link)
      cell.setCellValue(matchedInfo.name)
      cell.setCellStyle(style)

      cell = row.createCell(3)
      cell.setCellValue(matchedInfo.text)

      rowIndex += 1
    }

    Using.resource(new BufferedOutputStream(new FileOutputStream(file))) { outputStream =>
      workbook.write(outputStream)
    }
  }

  def r1c1ToA1(row: Int, col: Int): String = new CellReference(row, col).formatAsString()

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

  def search(path: String, regex: Regex): Array[MatchedInfo] = {
    val fileSearcher: FileVisitor[Path] = new FileVisitor[Path] {
      override def postVisitDirectory(dir: Path, exc: IOException) = FileVisitResult.CONTINUE

      override def preVisitDirectory(dir: Path, attrs: BasicFileAttributes) = FileVisitResult.CONTINUE

      override def visitFile(file: Path, attrs: BasicFileAttributes): FileVisitResult = {
        if (file.toFile.getName.matches(".*\\.xls[xm]?$")) search(file.toFile, regex)
        FileVisitResult.CONTINUE
      }

      override def visitFileFailed(file: Path, exc: IOException) = FileVisitResult.CONTINUE
    }

    Files.walkFileTree(Paths.get(path), fileSearcher)

    resultBuffer.toArray
  }

  private def search(file: File, regex: Regex): Unit = {
    Using.resource(WorkbookFactory.create(file, null, true)) { workbook =>
      val bookName = file.getName

      if (regex.findFirstIn(bookName).nonEmpty) {
        resultBuffer += new MatchedBookNameInfo(file, bookName)
      }

      workbook.forEach(sheet => {
        val sheetName = sheet.getSheetName

        if (regex.findFirstIn(sheetName).nonEmpty) {
          resultBuffer += new MatchedSheetNameInfo(file, sheetName, sheetName)
        }

        sheet.forEach(row => row.forEach(cell => {
          Try(getCellValue(cell)).foreach { value =>
            if (value.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(value).nonEmpty) {
              resultBuffer += new MatchedCellInfo(file, sheetName, cell.getRowIndex, cell.getColumnIndex, value)
            }
          }

          Try(cell.getCellComment.getString.getString).foreach { comment =>
            if (comment.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(comment).nonEmpty) {
              resultBuffer += new MatchedCommentInfo(file, sheetName, cell.getRowIndex, cell.getColumnIndex, comment)
            }
          }
        }))

        val drawingPatriarch = sheet.getDrawingPatriarch

        if (drawingPatriarch != null) {
          drawingPatriarch.forEach(shape => {
            def processShape(shapeName: String, row: Int, col: Int, value: String): Unit = {
              if (value.split("\r?\n").exists(s => regex.findFirstIn(s).nonEmpty) || regex.findFirstIn(value).nonEmpty) {
                resultBuffer += new MatchedShapeInfo(file, sheetName, shapeName, row, col, value)
              }
            }

            walkShape(shape, 0, 0, processShape)
          })
        }
      })
    }
  }

  private def walkShape(shape: Any, ancestorRow: Int, ancestorColumn: Int, processShape: (String, Int, Int, String) => ()): Unit = {
    def getRowColumnIndex(shape: Shape): (Int, Int) =
      shape.getAnchor match {
        case null => (ancestorRow, ancestorColumn)
        case anchor: XSSFClientAnchor => (anchor.getRow1, anchor.getCol1.toInt)
        case anchor: HSSFClientAnchor => (anchor.getRow1, anchor.getCol1.toInt)
        case _ => (ancestorRow, ancestorColumn)
      }

    shape match {
      case shape: XSSFSimpleShape =>
        val (row, col) = getRowColumnIndex(shape)
        processShape(shape.getShapeName, row, col, shape.getText)
      case shape: HSSFSimpleShape =>
        val (row, col) = getRowColumnIndex(shape)
        processShape(shape.getShapeName, row, col, shape.getString.getString)
      case group: XSSFShapeGroup =>
        val (row, col) = getRowColumnIndex(group)
        group.forEach(walkShape(_, row, col, processShape))
      case group: HSSFShapeGroup =>
        val (row, col) = getRowColumnIndex(group)
        group.forEach(walkShape(_, row, col, processShape))
      case _ =>
    }
  }

}
