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
    // Используем try-with-resources для автоматического закрытия InputStream
    InputStream inputStream = session.read(flowFile)
    Workbook wb = WorkbookFactory.create(inputStream)
    inputStream.close() // Закрываем InputStream

    def sheet = wb.getSheetAt(0) // Используем первый лист
    int lastRow = sheet.getLastRowNum()
    def non_empty_columns = []

    // Находим непустые столбцы
    for (def i = 0; i <= lastRow; i++) {
        def row = sheet.getRow(i)
        if (row == null) continue
        int lastColumn = row.getLastCellNum()
        for (def j = 0; j < lastColumn; j++) {
            Cell cell = row.getCell(j)
            if (cell != null && !(cell.getCellType() == CellType.BLANK)) {
                if (!non_empty_columns.contains(j)) {
                    non_empty_columns.add(j)
                }
            }
        }
    }

    // Записываем данные в новый FlowFile с использованием CSVWriter
    newFlowFile = session.write(newFlowFile, { outputStream ->
        OutputStreamWriter outputStreamWriter = new OutputStreamWriter(outputStream, StandardCharsets.UTF_8)
        // Создаем CSVWriter с настройками
        def csvWriter = new CSVWriter(outputStreamWriter,
            CSVWriter.DEFAULT_SEPARATOR,      // разделитель (запятая)
            CSVWriter.NO_QUOTE_CHARACTER,     // символ кавычки (отключен)
            CSVWriter.DEFAULT_ESCAPE_CHARACTER, // символ экранирования
            CSVWriter.DEFAULT_LINE_END        // конец строки
        )

        // Обрабатываем строки и записываем их в CSV
        for (def i = 0; i <= lastRow; i++) {
            def row = sheet.getRow(i)
            if (row == null) continue
            def rowData = []
            non_empty_columns.each { colIndex ->
                Cell cell = row.getCell(colIndex)
                def value = ""
                if (cell == null || cell.getCellType() == CellType.BLANK) {
                    value = ""
                } else if (cell.getCellType() == CellType.NUMERIC) {
                    if (DateUtil.isCellDateFormatted(cell)) {
                        Date date = cell.getDateCellValue()
                        SimpleDateFormat format1 = new SimpleDateFormat("dd-MM-yyyy")
                        value = format1.format(date)
                    } else {
                        double numericValue = cell.getNumericCellValue()
                        // Проверяем, является ли число целым
                        if (numericValue == Math.floor(numericValue)) {
                            value = String.valueOf((int) numericValue) // Преобразуем в целое число
                        } else {
                            value = String.valueOf(numericValue) // Оставляем как есть, если число дробное
                        }
                    }
                } else if (cell.getCellType() == CellType.BOOLEAN) {
                    value = cell.getBooleanCellValue().toString() // Преобразуем boolean в строку
                } else if (cell.getCellType() == CellType.FORMULA) {
                    value = cell.getCellFormula()
                } else {
                    value = cell.getStringCellValue().replace("\n", "")
                }
                rowData.add(value)
            }
            csvWriter.writeNext(rowData as String[])
        }

        csvWriter.close() // Закрываем CSVWriter
    } as OutputStreamCallback)

    // Закрываем Workbook
    wb.close()

    // Удаляем исходный FlowFile
    session.remove(flowFile)

    // Передаем новый FlowFile в REL_SUCCESS
    session.transfer(newFlowFile, REL_SUCCESS)
} catch (ex) {
    // В случае ошибки удаляем новый FlowFile и передаем исходный в REL_FAILURE
    session.remove(newFlowFile)
    log.error("Error processing Excel file: " + ex.getMessage())
    session.transfer(flowFile, REL_FAILURE)
}
