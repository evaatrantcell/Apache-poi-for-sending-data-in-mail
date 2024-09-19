@Grab(group='org.apache.poi', module='poi-ooxml', version='5.2.3')
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

def call(filePath, sheetName) {
    def f1 = new File(filePath)
    def testrunexcelfolder = f1.text

    // Get all .xlsx files in the directory
    def files = new File(directoryPath).listFiles().findAll { it.name.endsWith(".xlsx") }
    
    def tableHtml = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    files.each { file ->
        FileInputStream fis = new FileInputStream(file)
        XSSFWorkbook workbook = new XSSFWorkbook(fis)
        
        // Check if the sheet exists in the workbook
        if (workbook.getSheet(sheetName)) {
            println "Sheet '${sheetName}' found in file: ${file.getAbsolutePath()}"
            Sheet sheet = workbook.getSheet(sheetName)
            for (Row row : sheet) {
                tableHtml += "<tr>"
                for (Cell cell : row) {
                    String cellValue = ""
                    switch (cell.cellType) {
                        case CellType.STRING:
                            cellValue = cell.stringCellValue
                            break
                        case CellType.NUMERIC:
                            cellValue = cell.numericCellValue.toString()
                            break
                        case CellType.BOOLEAN:
                            cellValue = cell.booleanCellValue.toString()
                            break
                        default:
                            cellValue = ""
                    }
                    tableHtml += "<td>${cellValue}</td>"
                }
                tableHtml += "</tr>"
            }
            tableHtml += "</table>"
            }
        workbook.close()
        fis.close()
    }
    // FileInputStream file = new FileInputStream(new File(filePath))
    // Workbook workbook = new XSSFWorkbook(file)
    // Sheet sheet = workbook.getSheet(sheetName)
    // file.close()
    return tableHtml
}