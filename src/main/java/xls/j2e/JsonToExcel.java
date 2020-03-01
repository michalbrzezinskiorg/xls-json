package xls.j2e;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import lombok.var;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.BufferedOutputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.time.LocalDateTime;
import java.util.Date;
import java.util.Map;
import java.util.concurrent.atomic.AtomicInteger;
import java.util.concurrent.atomic.AtomicReference;

public class JsonToExcel {

    BufferedOutputStream getXls(String json, String fileName) {
        XSSFWorkbook workbook = new XSSFWorkbook();
        ObjectMapper mapper = new ObjectMapper();
        var map = getMap(json, mapper);
        map.entrySet().iterator().forEachRemaining(sheetData -> {
            AtomicReference<XSSFSheet> sheet = new AtomicReference<>();
            sheet.set(workbook.createSheet(sheetData.getKey()));
            AtomicInteger iterableRows = new AtomicInteger(0);
            sheetData.getValue().entrySet().iterator().forEachRemaining(rowData -> {
                Integer key = getKey(rowData.getKey(), iterableRows);
                Row row = sheet.get().createRow(key);
                AtomicInteger iterableCells = new AtomicInteger(0);
                rowData.getValue().entrySet().iterator().forEachRemaining(k -> {
                    Object o = k.getValue();
                    Cell cell = row.createCell(getKey(k.getKey(), iterableCells));
                    if (isInteger(o)) cell.setCellValue(Integer.parseInt(o.toString()));
                    else if (isDouble(o)) cell.setCellValue(Double.parseDouble(o.toString()));
                    else if (isDate(o)) cell.setCellValue((Date) o);
                    else if (isLocalDateTime(0)) cell.setCellValue((LocalDateTime) o);
                    else if (isBoolean()) cell.setCellValue((boolean) o);
                    else if (o!=null) cell.setCellValue((String) o);
                    else cell.setCellValue("null");
                });
            });
        });
        return getBufferedOutputStream(fileName, workbook);
    }

    private boolean isBoolean() {
        return false;
    }

    private boolean isLocalDateTime(int i) {
        return false;
    }

    private boolean isDate(Object o) {
        return false;
    }

    private boolean isInteger(Object o) {
        try{
            String s = String.valueOf(o).trim();
            int i = Integer.valueOf(s);
            String s1 = String.valueOf(i);
            if(s.length()==s1.length())
                return true;
        } catch (RuntimeException e) {
        }
        return false;
    }

    private boolean isDouble(Object o) {
        try{
            String s = o.toString().trim();
            double i = Double.parseDouble(s);
            String s1 = String.valueOf(i);
            if(s.length()==s1.length())
                return true;
        } catch (RuntimeException e) {
        }
        return false;
    }

    private int getKey(String key, AtomicInteger i) {
        try{
            i.incrementAndGet();
            return Integer.parseInt(key);
        } catch (RuntimeException e){
            return i.get();
        }
    }

    private BufferedOutputStream getBufferedOutputStream(String fileName, XSSFWorkbook workbook) {
        try {
            FileOutputStream fos = new FileOutputStream(fileName);
            BufferedOutputStream bos = new BufferedOutputStream(fos);
            workbook.write(fos);
            return bos;
        } catch (IOException e) {
            e.printStackTrace();
            throw new IllegalArgumentException(e.getClass() + " ][ " + e.getMessage());
        }
    }

    private Map<String, Map<String, Map<String, Object>>> getMap(String json, ObjectMapper mapper) {
        Map<String, Map<String, Map<String, Object>>> map;
        try {
            return map = mapper.readValue(json, Map.class);
        } catch (JsonProcessingException e) {
            throw new IllegalArgumentException
                    ("wrong format: expected Map<String sheetname, Map<String tablename, Map<Integer rownumber, Object>>> " +
                            "where legit Object can be: boolean, Date, LocalDateTime, String, int, double");
        }
    }
}
