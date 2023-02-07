package pe.fwani.convert.utils;

import org.apache.poi.ss.usermodel.Sheet;

import java.util.ArrayList;
import java.util.List;

public class Functions {
    public static List<List<String>> sheetToList(Sheet sheet) {
        List<List<String>> expectRows = new ArrayList<>();
        for (var row : sheet) {
            List<String> expectValues = new ArrayList<>();
            for (var cell : row) {
                expectValues.add(cell.getStringCellValue());
            }
            expectRows.add(expectValues);
        }
        return expectRows;
    }
}
