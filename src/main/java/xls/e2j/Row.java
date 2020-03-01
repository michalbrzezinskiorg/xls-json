package xls.e2j;

import lombok.Data;

import java.util.HashMap;
import java.util.Map;

@Data
public final class Row {
    public Map<Integer, String> columns = new HashMap<>();
    public Row(Map<Integer, String> value) {
        columns = value;
    }
}
