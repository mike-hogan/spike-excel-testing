package spike

import org.apache.poi.ss.util.CellReference
import org.apache.poi.xssf.usermodel.XSSFCell
import org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator
import org.apache.poi.xssf.usermodel.XSSFSheet
import org.apache.poi.xssf.usermodel.XSSFWorkbook

class TestSheet(val path: String) {
    private var evaluator: XSSFFormulaEvaluator
    private var sheet: XSSFSheet

    init {
        val wb = XSSFWorkbook(javaClass.getResourceAsStream(path))
        this.sheet = wb.getSheetAt(0)
        this.evaluator = wb.creationHelper.createFormulaEvaluator()
    }

    fun number(ref: String): Number {
        val cell = getCell(ref)
        evaluator.evaluateFormulaCell(cell)
        return evaluator.evaluate(cell).numberValue
    }


    private fun getCell(ref: String): XSSFCell {
        val cellReference = CellReference(ref)
        return sheet
                .getRow(cellReference.row)
                .getCell(cellReference.col.toInt())
    }

    fun set(ref: String, n: Number) {
        getCell(ref).setCellValue(n.toDouble())
    }
}
