package spike

import io.kotest.matchers.shouldBe
import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.TestInstance

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class SpikeTest {

    @Test
    fun `the spreadsheet defaults to 6`() {
        val (sheet, evaluator) = sheet()
        val d2 = getCell(sheet, "D2")

        val cellValue = evaluator.evaluate(d2)
        cellValue.formatAsString().shouldBe("6.0")
    }

    @Test
    fun `7 plus 9 = 16`() {
        val (sheet, evaluator) = sheet()
        val b2 = getCell(sheet, "B2")
        val c2 = getCell(sheet, "C2")
        val d2 = getCell(sheet, "D2")

        b2.setCellValue(7.0)
        c2.setCellValue(9.0)
        evaluator.evaluateFormulaCell(d2)
        val cellValue = evaluator.evaluate(d2)
        cellValue.formatAsString().shouldBe("16.0")
    }

    private fun getCell(sheet: XSSFSheet, ref: String): XSSFCell {
        val cellReference = CellReference(ref)
        return sheet
                .getRow(cellReference.row)
                .getCell(cellReference.col.toInt())
    }

    private fun sheet(): Pair<XSSFSheet, XSSFFormulaEvaluator> {
        val wb = XSSFWorkbook(javaClass.getResourceAsStream("/SimpleMath.xlsx"))
        val sheet = wb.getSheetAt(0)
        val evaluator = wb.creationHelper.createFormulaEvaluator()
        return Pair(sheet, evaluator)
    }
}
