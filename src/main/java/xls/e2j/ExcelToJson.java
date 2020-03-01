package xls.e2j;

import com.fasterxml.jackson.core.JsonProcessingException;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.poifs.filesystem.FileMagic;
import org.json.JSONObject;

import java.io.BufferedInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.HashMap;
import java.util.Map;
import java.util.stream.Collectors;

public class ExcelToJson {

    private ThreadLocal<Map<String, Map<Integer, Map<Integer, String>>>> sheetList = new ThreadLocal();
    private ObjectMapper objectMapper = new ObjectMapper();

    public Map<String, Map<Integer, Map<Integer, String>>> getMap(BufferedInputStream inputStream) {
        return process(inputStream);
    }

    public Map<String, Table> getTables(BufferedInputStream inputStream) {
        return makeTable(process(inputStream));
    }

    public String getJsonString(BufferedInputStream inputStream) {
        try {
            return objectMapper.writeValueAsString(process(inputStream));
        } catch (JsonProcessingException e) {
            throw new RuntimeException(e.getMessage());
        }
    }

    private Map<String, Map<Integer, Map<Integer, String>>> process(InputStream inputStream) {
        Props props = new Props(inputStream);
        sheetList.set(new HashMap<>());
        XlsProcessor xlsProcessor = new XlsProcessor();
        try {
            xlsProcessor.determineAndRunProcessing(props, sheetList.get());
        } catch (Exception e) {
            throw new IllegalArgumentException(e.getMessage());
        }
        return sheetList.get();
    }

    private Map<String, Table> makeTable(Map<String, Map<Integer, Map<Integer, String>>> sheetList) {
        return sheetList.entrySet().stream()
                .map(e -> new Table(e))
                .collect(Collectors.toMap(x -> x.getName(), x -> x));
    }

    static class Props {
        public final InputStream file;
        private boolean isxls;
        private boolean isxlsx;

        public Props(InputStream file) {
            this.file = file;
            isxls = isXLS(file);
            isxlsx = isXLSX(file);
        }

        public boolean isIsxls() {
            return isxls;
        }

        public boolean isXlsx() {
            return isxlsx;
        }

        private boolean isXLS(InputStream file) {
            try {
                return FileMagic.valueOf(file).equals(FileMagic.OLE2);
            } catch (IOException e) {
                return false;
            }
        }

        private boolean isXLSX(InputStream file) {
            try {
                return FileMagic.valueOf(file).equals(FileMagic.OOXML);
            } catch (IOException e) {
                return false;
            }
        }
    }
}
