package pe.fwani.convert;

import lombok.Getter;
import lombok.NonNull;
import lombok.Setter;
import org.apache.commons.csv.CSVFormat;
import org.apache.commons.math3.util.Pair;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStreamReader;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.atomic.AtomicInteger;

@Getter
@Setter
public class CsvToExcel {
    private String source;
    private String destination;
    private final CSVFormat csvFormat;
    private final int MAX_ROWS_OF_SHEET;

    private final boolean caseSensitive;
    private WorkbookType workbookType = WorkbookType.XSSF;


    public CsvToExcel(@NonNull String source, @NonNull String destination, @NonNull CSVFormat csvFormat, int MAX_ROWS_OF_SHEET, boolean caseSensitive) {
        this.source = source;
        this.destination = destination;
        this.csvFormat = csvFormat;
        this.MAX_ROWS_OF_SHEET = MAX_ROWS_OF_SHEET;
        this.caseSensitive = caseSensitive;
    }

    public CsvToExcel(@NonNull String source, @NonNull String destination, @NonNull CSVFormat csvFormat) {
        this(source, destination, csvFormat, 1048576, false);
    }

    public CsvToExcel(@NonNull String source, @NonNull String destination) {
        this(source, destination, CSVFormat.DEFAULT);
    }

    public void convert(@NonNull List<Pair<String, List<String>>> sheetNameAndColumnNamesPairs) throws IOException {
        convertEachSheetWithColumns(sheetNameAndColumnNamesPairs, false);
    }

    public void convert(@NonNull List<Pair<String, List<String>>> sheetNameAndColumnNamesPairs, boolean dropAllNoneRow) throws IOException {
        convertEachSheetWithColumns(sheetNameAndColumnNamesPairs, dropAllNoneRow);
    }

    public void convert() throws IOException {
        convertAll("Sheet1", false);
    }

    public void convert(@NonNull String sheetName) throws IOException {
        convertAll(sheetName, false);
    }

    public void convert(boolean dropAllNoneRow) throws IOException {
        convertAll("Sheet1", dropAllNoneRow);
    }

    public void convert(@NonNull String sheetName, boolean dropAllNoneRow) throws IOException {
        convertAll(sheetName, dropAllNoneRow);
    }

    private Workbook getWorkbook(WorkbookType workbookType) {
        Workbook workbook;
        if (workbookType.equals(WorkbookType.HSSF)) {
            workbook = new HSSFWorkbook();
        } else if (workbookType.equals(WorkbookType.XSSF)) {
            workbook = new XSSFWorkbook();
        } else { // SXSSF
            workbook = new SXSSFWorkbook();
        }
        return workbook;
    }

    private String getCaseSensitiveString(String origin) {
        return caseSensitive ? origin.toUpperCase() : origin;
    }

    private boolean addRow(Sheet sheet, int rowIndex, List<String> data, boolean dropAllNoneRow) {
        if (dropAllNoneRow && data.stream().allMatch(x -> x.equals(""))) {
            return false;
        }
        if (data.stream().allMatch(x -> x.equals(""))) {
            return true;
        }
        var row = sheet.createRow(rowIndex);
        var idx = new AtomicInteger();
        data.forEach(x -> {
            var cellId = idx.getAndIncrement();
            if (x.equals(""))
                return;
            row.createCell(cellId).setCellValue(x);
        });
        return true;
    }

    // multi sheet
    private void convertEachSheetWithColumns(List<Pair<String, List<String>>> sheetNameAndColumnNamesPairs,
                                             boolean dropAllRow
    ) throws IOException {
        try (var wb = getWorkbook(workbookType);
             var fileInputStream = new FileInputStream(source);
             var reader = new InputStreamReader(fileInputStream, StandardCharsets.UTF_8);
             var writer = new FileOutputStream(destination)
        ) {
            List<Pair<Sheet, List<Integer>>> sheetAndIndexes = new ArrayList<>();
            List<Integer> rowIdxList = new ArrayList<>();

            var parser = csvFormat.parse(reader);
            boolean isHeader = true;
            for (var record : parser) {
                if (isHeader) {  // header
                    isHeader = false;
                    var header = record.toList().stream()
                            .map(this::getCaseSensitiveString)
                            .toList();
                    for (var sheetNameAndColumnNames : sheetNameAndColumnNamesPairs) {
                        var sheet = wb.createSheet(sheetNameAndColumnNames.getKey());
                        var selectedIndexes = sheetNameAndColumnNames.getValue()
                                .stream().map(x -> header.indexOf(getCaseSensitiveString(x)))
                                .filter(x -> x >= 0)
                                .toList();
                        addRow(sheet, 0, selectedIndexes.stream().map(header::get).toList(), dropAllRow);
                        sheetAndIndexes.add(new Pair<>(sheet, selectedIndexes));
                        rowIdxList.add(1);
                    }
                } else {  // row
                    if (rowIdxList.stream().allMatch(x -> x == MAX_ROWS_OF_SHEET || x == -1)) {
                        break;
                    }
                    for (var i = 0; i < sheetAndIndexes.size(); i++) {
                        var rowIdx = rowIdxList.get(i);
                        if (rowIdx == MAX_ROWS_OF_SHEET) {
                            continue;
                        }
                        var pair = sheetAndIndexes.get(i);
                        if (addRow(pair.getKey(), rowIdx, pair.getValue().stream().map(record::get).toList(), dropAllRow)) {
                            rowIdxList.set(i, rowIdx + 1);
                        }
                    }
                }
            }
            wb.write(writer);
        }
    }

    // one sheet
    private void convertAll(String sheetName, boolean dropAllRow) throws IOException { // with header
        try (var wb = getWorkbook(workbookType);
             var fileInputStream = new FileInputStream(source);
             var reader = new InputStreamReader(fileInputStream, StandardCharsets.UTF_8);
             var writer = new FileOutputStream(destination)
        ) {
            var parser = csvFormat.parse(reader);
            var sheet = wb.createSheet(sheetName);
            int nowRowIdx = 0;
            for (var record : parser) {
                if (nowRowIdx == MAX_ROWS_OF_SHEET) {
                    break;
                }
                addRow(sheet, nowRowIdx++, record.stream().toList(), dropAllRow);
            }
            wb.write(writer);
        }
    }


}
