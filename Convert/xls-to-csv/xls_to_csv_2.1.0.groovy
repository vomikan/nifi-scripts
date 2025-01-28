@Grab(group='org.apache.poi', module='poi-ooxml', version='5.0.0')
@Grab(group='org.apache.poi', module='poi-ooxml-schemas', version='4.1.2')
@Grab(group='org.apache.xmlbeans', module='xmlbeans', version='3.1.0')
@Grab(group='commons-io', module='commons-io', version='2.8.0')
@Grab(group='com.opencsv', module='opencsv', version='5.5.2') // Add OpenCSV dependency
import org.apache.poi.ss.usermodel.*
import org.apache.poi.xssf.usermodel.XSSFWorkbook
import org.apache.poi.ss.util.*
import java.io.*
import java.text.SimpleDateFormat
import org.apache.commons.io.IOUtils
import java.nio.charset.StandardCharsets
import com.opencsv.CSVWriter // Import CSVWriter

flowFile = session.get()
if (!flowFile) return

newFlowFile = session.create()
def filename = flowFile.getAttribute('filename')
newFlowFile = session.putAttribute(newFlowFile, "filename", filename)
// Устанавливаем mimetype для CSV
newFlowFile = session.putAttribute(newFlowFile, "mime.type", "text/csv")

flowFile.getAttributes().each { key, value ->
    session.putAttribute(newFlowFile, key, value)
}

try {
    // Используем try-with-resources для автоматического закрытия InputStream и Workbook
    try (InputStream inputStream = session.read(flowFile);
         Workbook wb = WorkbookFactory.create(inputStream)) {

        def sheet = wb.getSheetAt(0) // Используем первый лист
        int lastRow = sheet.getLastRowNum()

        // Находим непустые столбцы
        Set<Integer> nonEmptyColumns = findNonEmptyColumns(sheet, lastRow)

        // Записываем данные в новый FlowFile с использованием CSVWriter
        newFlowFile = session.write(newFlowFile, { outputStream ->
            try (OutputStreamWriter outputStreamWriter = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8);
                 CSVWriter csvWriter = new CSVWriter(outputStreamWriter,
                     CSVWriter.DEFAULT_SEPARATOR,      // разделитель (запятая)
                     CSVWriter.NO_QUOTE_CHARACTER,     // символ кавычки (отключен)
                     CSVWriter.DEFAULT_ESCAPE_CHARACTER, // символ экранирования
                     CSVWriter.DEFAULT_LINE_END        // конец строки
                 )) {

                // Обрабатываем строки и записываем их в CSV
                for (def i = 0; i <= lastRow; i++) {
                    def row = sheet.getRow(i)
                    if (row == null) continue

                    def rowData = nonEmptyColumns.collect { colIndex ->
                        Cell cell = row.getCell(colIndex)
                        getCellValue(cell)
                    }

                    csvWriter.writeNext(rowData as String[])
                }
            }
        } as OutputStreamCallback)
    }

    // Удаляем исходный FlowFile
    session.remove(flowFile)
    // Передаем новый FlowFile в REL_SUCCESS
    session.transfer(newFlowFile, REL_SUCCESS)
} catch (Exception ex) {
    // В случае ошибки удаляем новый FlowFile и передаем исходный в REL_FAILURE
    session.remove(newFlowFile)
    log.error("Error processing Excel file: " + ex.getMessage(), ex)
    session.transfer(flowFile, REL_FAILURE)
}

// Метод для поиска непустых столбцов
Set<Integer> findNonEmptyColumns(Sheet sheet, int lastRow) {
    Set<Integer> nonEmptyColumns = new HashSet<>()
    for (def i = 0; i <= lastRow; i++) {
        def row = sheet.getRow(i)
        if (row == null) continue
        int lastColumn = row.getLastCellNum()
        for (def j = 0; j < lastColumn; j++) {
            Cell cell = row.getCell(j)
            if (cell != null && cell.cellType != CellType.BLANK) {
                nonEmptyColumns.add(j)
            }
        }
    }
    return nonEmptyColumns
}

// Метод для получения значения ячейки
String getCellValue(Cell cell) {
    if (cell == null || cell.cellType == CellType.BLANK) {
        return ""
    } else if (cell.cellType == CellType.NUMERIC) {
        if (DateUtil.isCellDateFormatted(cell)) {
            Date date = cell.dateCellValue
            SimpleDateFormat format1 = new SimpleDateFormat("dd-MM-yyyy")
            return format1.format(date)
        } else {
            double numericValue = cell.numericCellValue
            // Проверяем, является ли число целым
            if (numericValue == Math.floor(numericValue)) {
                return String.valueOf((int) numericValue) // Преобразуем в целое число
            } else {
                return String.valueOf(numericValue) // Оставляем как есть, если число дробное
            }
        }
    } else if (cell.cellType == CellType.BOOLEAN) {
        return cell.booleanCellValue.toString() // Преобразуем boolean в строку
    } else if (cell.cellType == CellType.FORMULA) {
        return cell.cellFormula
    } else {
        return cell.stringCellValue.replace("\n", "")
    }
}
