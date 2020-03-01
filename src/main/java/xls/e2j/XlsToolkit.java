package xls.e2j;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.ReadOnlySharedStringsTable;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler.SheetContentsHandler;
import org.apache.poi.xssf.model.StylesTable;
import org.apache.poi.xssf.usermodel.XSSFComment;
import org.xml.sax.ContentHandler;
import org.xml.sax.InputSource;
import org.xml.sax.SAXException;
import org.xml.sax.XMLReader;

import javax.xml.parsers.ParserConfigurationException;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.logging.Level;
import java.util.logging.Logger;

final class XlsToolkit {

    private final TableMapper output;
    private OPCPackage xlsxPackage;

    XlsToolkit(InputStream inputStream) {
        output = new TableMapper();
        try {
            xlsxPackage = OPCPackage.open(inputStream);
        } catch (IOException | InvalidFormatException e) {
            throw new IllegalArgumentException("file format is not supported");
        }
    }

    private void processSheet(StylesTable styles, ReadOnlySharedStringsTable strings, SheetContentsHandler sheetHandler,
                              InputStream sheetInputStream) {
        DataFormatter formatter = new DataFormatter();
        InputSource sheetSource = new InputSource(sheetInputStream);
        try {
            XMLReader sheetParser = XMLHelper.newXMLReader();
            ContentHandler handler = new XSSFSheetXMLHandler(styles, null, strings, sheetHandler, formatter, false);
            sheetParser.setContentHandler(handler);
            sheetParser.parse(sheetSource);
        } catch (ParserConfigurationException | SAXException | IOException e) {
            Logger.getAnonymousLogger().log(Level.WARNING, e.getMessage());
        }
    }

    public Map<String, Map<Integer, Map<Integer, String>>> process() throws IOException, OpenXML4JException, SAXException {
        Map<String, Map<Integer, Map<Integer, String>>> res = new HashMap<>();
        ReadOnlySharedStringsTable strings = new ReadOnlySharedStringsTable(this.xlsxPackage);
        XSSFReader xssfReader = new XSSFReader(this.xlsxPackage);
        StylesTable styles = xssfReader.getStylesTable();
        XSSFReader.SheetIterator iter = (XSSFReader.SheetIterator) xssfReader.getSheetsData();
        while (iter.hasNext()) {
            InputStream stream = iter.next();
            processSheet(styles, strings, new Harvester(), stream);
            stream.close();
            Map<Integer, Map<Integer, String>> sheet = output.getSheet();
            res.put(iter.getSheetName(), sheet);
        }
        return res;
    }

    private class Harvester implements SheetContentsHandler {

        @Override
        public void startRow(int rowNum) {
            output.newRow();
        }

        @Override
        public void endRow(int rowNum) {
            output.appendRow(rowNum);
        }

        @Override
        public void cell(String cellReference, String formattedValue, XSSFComment comment) {
            output.appendCell((new CellReference(cellReference)).getCol(), formattedValue);
        }

        @Override
        public void headerFooter(String s, boolean b, String s1) {
            // posible implementation ;)
        }
    }
}
