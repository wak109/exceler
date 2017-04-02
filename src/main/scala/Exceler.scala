import org.apache.poi.ss.usermodel._

import java.io._

object Exceler {

    def excel(filename:String) : Unit  = {
        val workbook = WorkbookFactory.create(new File(filename))
        val sheet = workbook.getSheet("test")
        val row = sheet.getRow(0)
        val cell = row.getCell(0)
        val value = cell.getCellType match {
            case Cell.CELL_TYPE_BLANK => cell.getStringCellValue
            case Cell.CELL_TYPE_BOOLEAN => cell.getBooleanCellValue
            case Cell.CELL_TYPE_ERROR => cell.getErrorCellValue
            case Cell.CELL_TYPE_FORMULA => cell.getCellFormula
            case Cell.CELL_TYPE_NUMERIC => cell.getNumericCellValue
            case Cell.CELL_TYPE_STRING => cell.getStringCellValue
            }
        println(value)
    }
}
