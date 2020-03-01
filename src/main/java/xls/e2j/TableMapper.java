package xls.e2j;

import java.util.HashMap;
import java.util.Map;

final class TableMapper {

    private Map<Integer, String> line = new HashMap<>();
    private final Map<Integer, Map<Integer, String>> sheet = new HashMap<>();

    TableMapper() { }

    public void appendCell(int index, String value) {
        line.put(index, value);
    }

    public void newRow() {
        line = new HashMap<>();
    }

    public void appendRow(int index) {
        if(!line.isEmpty()) sheet.put(index, line);
    }

    public Map<Integer, Map<Integer, String>> getSheet() {
        return new HashMap<>(sheet);
    }
}
