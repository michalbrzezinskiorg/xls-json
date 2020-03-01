package xls.e2j;

import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.OpenXML4JException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.*;
import org.xml.sax.SAXException;

import java.io.IOException;
import java.util.HashMap;
import java.util.Map;

final class XlsProcessor {

    void determineAndRunProcessing(ExcelToJson.Props s, Map<String, Map<Integer, Map<Integer, String>>> sheetList) throws IOException, OpenXML4JException, SAXException {
        boolean isxls = s.isIsxls();
        boolean xlsx = s.isXlsx();
        if (!isxls && !xlsx) {
            throw new IllegalArgumentException();
        } else if (xlsx) {
            processXlsx(s, sheetList);
        } else {
            processXls(s, sheetList);
        }
    }

    void processXlsx(ExcelToJson.Props s, Map<String, Map<Integer, Map<Integer, String>>> sheetList) throws IOException, OpenXML4JException, SAXException {
        XlsToolkit xlsx = new XlsToolkit(s.file);
        xlsx.process().entrySet().forEach(e -> sheetList.put(e.getKey(), e.getValue()));
    }

    void processXls(ExcelToJson.Props props, Map<String, Map<Integer, Map<Integer, String>>> sheetList) throws IOException, InvalidFormatException {
        Workbook wb = WorkbookFactory.create(props.file);
        int sheets = wb.getNumberOfSheets();
        processSheets(sheetList, wb, sheets);
    }

    void processSheets(Map<String, Map<Integer, Map<Integer, String>>> sheetList, Workbook wb, int sheets) {
        for (int i = 0; i < sheets; i++) {
            Sheet mySheet = wb.getSheetAt(i);
            mySheet.setDisplayFormulas(true);
            String name = mySheet.getSheetName();
            Map recentsheet = new HashMap();
            processRows(mySheet, recentsheet);
            sheetList.put(name, recentsheet);
        }
    }

    void processRows(Sheet mySheet, Map recentSheet) {
        for (Row row : mySheet) {
            Map rowList = new HashMap();
            processCell(row, rowList);
            if(rowList.entrySet().size()>0)
            recentSheet.put(row.getRowNum(), rowList);
        }
    }

    void processCell(Row row, Map rowList) {
        for (Cell cell : row) {
            CellType type = cell.getCellType();
            String val = resolveTypeOfCell(cell, type);
            if(val!=null)
                rowList.put(cell.getColumnIndex(), val);
        }
    }

    String resolveTypeOfCell(Cell cell, CellType type) {
        String val;
        try {
            switch (type) {
                case BOOLEAN:
                    val = String.valueOf(cell.getBooleanCellValue());
                    break;
                case FORMULA:
                    val = resolveTypeOfCell(cell, cell.getCachedFormulaResultType());
                    break;
                case STRING:
                    val = cell.getStringCellValue();
                    break;
                case NUMERIC:
                    val = determineNumericType(cell);
                    break;
                default:
                    val = null;
                    break;
            }
        } catch (Exception e) {
            val = "error";
        }
        return val;
    }

    private String determineNumericType(Cell cell) {
        String val;
        double d = cell.getNumericCellValue();
        int i = (int) d;
        if(i == d)
            val = String.valueOf(i);
        else
            val = String.valueOf(d);
        return val;
    }
}