@Grab(group='org.apache.poi', module='poi-ooxml', version='5.2.3')
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import java.io.File
import java.io.FileInputStream

def call(filePath, sheetName, excelfolder) {
    def f1 = new File(filePath)
    def testrunexcelfolder = f1.text
    def excel_report_file_path = excelfolder + testrunexcelfolder
    // Get all .xlsx files in the directory
    def files = new File(excel_report_file_path).listFiles().findAll { it.name.endsWith(".xlsx") }
    def tableHtml = "<table border='1' cellpadding='5' cellspacing='0' style='border-collapse: collapse;'>"
    files.each { file ->
        FileInputStream fis = new FileInputStream(file)
        XSSFWorkbook workbook = new XSSFWorkbook(fis)
        String fileName = file.getName();
        // Check if the sheet exists in the workbook
        if (workbook.getSheet(sheetName)) {
            println "Sheet '${sheetName}' found in file: ${file.getAbsolutePath()}"
            // List to store file paths
            Sheet sheet = workbook.getSheet(sheetName);
            boolean isHeader = true; // Flag to identify the header row
            tableHtml += "<tr>" + "<th style='background-color: #FFD700;'>" + fileName + "</th>" + "</tr>";
            for (Row row : sheet) {
                tableHtml += "<tr>";
                for (Cell cell : row) {
                    String cellValue = "";
                    switch (cell.getCellType()) {
                        case CellType.STRING:
                            cellValue = cell.getStringCellValue();
                            break;
                        case CellType.NUMERIC:
                            cellValue = String.valueOf(cell.getNumericCellValue());
                            break;
                        case CellType.BOOLEAN:
                            cellValue = String.valueOf(cell.getBooleanCellValue());
                            break;
                        default:
                            cellValue = "";
                    }

                    // Add style for the header row
                    if (isHeader) {
                        tableHtml += "<th style='background-color: yellow;'>" + cellValue + "</th>";
                    } else {
                        String cellStyle = ""; // Variable to store background color
                        switch (cellValue.toUpperCase()) {
                            case "PASSED":
                                cellStyle = "background-color: lightgreen;";
                                break;
                            case "FAILED":
                                cellStyle = "background-color: lightcoral;"; // Light red
                                break;
                            case "SKIPPED":
                                cellStyle = "background-color: #FFA590;"; // Custom light orange
                                break;
                            case "WARNING":
                                cellStyle = "background-color: #FFFF98;"; // Light yellow
                                break;
                            default:
                                cellStyle = ""; // No style for other values
                        }
                        tableHtml += "<td style='" + cellStyle + "'>" + cellValue + "</td>";
                    }
                }
                tableHtml += "</tr>";

                // After processing the first row, set the flag to false
                isHeader = false;
            }

        }
        workbook.close()
        fis.close()
    }
    tableHtml += "</table>"
    return tableHtml
}
