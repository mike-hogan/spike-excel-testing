package spike

import io.kotest.matchers.shouldBe
import org.junit.jupiter.api.Test
import org.junit.jupiter.api.TestInstance

@TestInstance(TestInstance.Lifecycle.PER_CLASS)
class SpikeTest {

    @Test
    fun `the spreadsheet defaults to 6`() {
        val sheet = TestSheet("/SimpleMath.xlsx")

        sheet.number("D2").shouldBe(6)
    }

    @Test
    fun `7 plus 9 = 16`() {
        val sheet = TestSheet("/SimpleMath.xlsx")

        sheet.set("B2", 7)
        sheet.set("C2", 9)

        sheet.number("D2").shouldBe(16)
    }
}
