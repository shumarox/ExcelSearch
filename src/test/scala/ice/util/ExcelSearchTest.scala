package ice.util

import org.junit.Test

class ExcelSearchTest {
  @Test
  def mainTest(): Unit = {
    ExcelSearch.main(Array[String]("src/test/resources", ".", "src/test/resources/result.xlsx"))
  }
}
