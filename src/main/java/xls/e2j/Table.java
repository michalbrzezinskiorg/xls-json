package xls.e2j;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;
import java.util.Set;

@Data
public final class Table {

    private final String name;
    private Map<Integer, Row> rows = new HashMap<>();

    public Table(Map.Entry<String, Map<Integer, Map<Integer, String>>> e) {
        name = e.getKey();
        Map<Integer, Map<Integer, String>> value = e.getValue();
        Set<Map.Entry<Integer, Map<Integer, String>>> entry = value.entrySet();
        entry.forEach(en -> addRow(en));
    }

    private void addRow(Map.Entry<Integer, Map<Integer, String>> entries) {
        rows.put(entries.getKey(), new Row(entries.getValue()));
    }

}
