import org.apache.commons.csv.CSVFormat
import org.apache.commons.csv.CSVPrinter
import org.apache.poi.openxml4j.opc.OPCPackage
import org.apache.poi.openxml4j.opc.PackageAccess
import org.apache.poi.ss.usermodel.DataFormatter
import org.apache.poi.util.XMLHelper
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable
import org.apache.poi.xssf.eventusermodel.XSSFReader
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler
import org.apache.poi.xssf.usermodel.XSSFComment
import org.xml.sax.InputSource
import java.io.BufferedWriter
import java.io.File
import java.io.OutputStreamWriter

/**
 * example of streaming read of .xlsx file using Apache POI
 *
 * see also:
 * - https://poi.apache.org/components/spreadsheet/how-to.html#xssf_sax_api
 * - https://poi.apache.org/components/spreadsheet/examples.html#xssf-only
 * - https://svn.apache.org/repos/asf/poi/trunk/poi-examples/src/main/java/org/apache/poi/examples/xssf/eventusermodel/XLSX2CSV.java
 */
fun main(args: Array<String>) {

    if (args.isEmpty()) {
        System.err.println("Usage: java -jar build/libs/poi-kotlin-example-1.0-SNAPSHOT.jar <.xlsx>")
        return
    }

    OPCPackage.open(File(args[0]).path, PackageAccess.READ).use { xlsxPackage ->
        val printer = CSVPrinter(BufferedWriter(OutputStreamWriter(System.out)), CSVFormat.DEFAULT)
        val xssfReader = XSSFReader(xlsxPackage)
        val sheetIterator = xssfReader.sheetsData as XSSFReader.SheetIterator
        for (inputStream in sheetIterator) {
            printer.printRecord(sheetIterator.sheetName)
            XMLHelper.newXMLReader().apply {
                contentHandler = XSSFSheetXMLHandler(
                    xssfReader.getStylesTable(),
                    null,
                    ReadOnlySharedStringsTable(xlsxPackage),
                    object : XSSFSheetXMLHandler.SheetContentsHandler {
                        override fun startRow(rowNum: Int) = Unit
                        override fun endRow(rowNum: Int) = printer.println()
                        override fun cell(cellReference: String?, formattedValue: String?, comment: XSSFComment?) = printer.print(formattedValue)
                    },
                    DataFormatter(true),
                    false
                )
            }.parse(InputSource(inputStream))
        }
        printer.flush()
    }
}
