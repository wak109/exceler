/* vim: set ts=4 et sw=4 sts=4 fileencoding=utf-8: */
import org.apache.poi.ss.usermodel._
import org.apache.poi.xssf.usermodel.XSSFWorkbook

import java.io._
import java.nio.file._

object FileOp {

    def createDirectories(filename:String) : Path = {
        val dir = Paths.get(filename).getParent()
        if (dir != null)
            Files.createDirectories(dir)
        else
            null
    }

    def exists(filename:String) : Boolean = {
        Files.exists(Paths.get(filename))
    }
}

object WorkbookOp {

    def open(filename:String): Workbook =  {
        WorkbookFactory.create(new File(filename))
    }

    def create(filename:String): Workbook =  {
        new XSSFWorkbook()
    }

    def save(workbook:Workbook, filename:String): Unit =  {
        FileOp.createDirectories(filename)
        val out = new FileOutputStream(filename)
        workbook.write(out)
        out.close()
    }
}

object Exceler {

    def excel(filename:String) : Unit  = {
        // val workbook = WorkbookFactory.create(new File(filename))
        val workbook = WorkbookOp.open(filename)
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
        WorkbookOp.save(workbook, "test2.xlsx")
    }
}
