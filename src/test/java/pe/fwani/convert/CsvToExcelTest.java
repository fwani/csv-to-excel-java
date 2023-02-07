package pe.fwani.convert;

import org.apache.commons.math3.util.Pair;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.lang.reflect.InvocationTargetException;
import java.util.List;
import java.util.Objects;

import static org.assertj.core.api.Assertions.assertThat;
import static pe.fwani.convert.utils.Functions.sheetToList;

class CsvToExcelTest {
    private String origin;

    @BeforeEach
    void setUp() {
        origin = Objects.requireNonNull(getClass().getResource("/example/origin.csv")).getPath();
    }

    private void checkActualAndExpect(File actual, String expectPath, WorkbookType workbookType) throws IOException, ClassNotFoundException, NoSuchMethodException, InvocationTargetException, InstantiationException, IllegalAccessException {
        String workbookClassName;
        if (workbookType.equals(WorkbookType.HSSF)) {
            workbookClassName = HSSFWorkbook.class.getName();
        } else if (workbookType.equals(WorkbookType.XSSF)) {
            workbookClassName = XSSFWorkbook.class.getName();
        } else {
            workbookClassName = XSSFWorkbook.class.getName();
        }
        try (var expectReader = getClass().getResourceAsStream(expectPath);
             var actualReader = new FileInputStream(actual)) {
            assert expectReader != null;
            try (var expectWB = ((Workbook) Class.forName(workbookClassName)
                    .getConstructor(InputStream.class)
                    .newInstance(expectReader));
                 var actualWB = (Workbook) Class.forName(workbookClassName)
                         .getConstructor(InputStream.class)
                         .newInstance(actualReader)
            ) {
                assertThat(sheetToList(actualWB.getSheet("Sheet1")))
                        .isEqualTo(sheetToList(expectWB.getSheet("Sheet1")));
            }
        }
    }

    @Test
    void hssfTest() throws IOException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        var dest = File.createTempFile("dest-", ".xls");
        dest.deleteOnExit();

        var convertor = new CsvToExcel(origin, dest.getAbsolutePath());
        convertor.setWorkbookType(WorkbookType.HSSF);
        convertor.convert();

        checkActualAndExpect(dest, "/example/result-all.xls", WorkbookType.HSSF);
    }

    @Test
    void xssfTest() throws IOException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        var dest = File.createTempFile("dest-", ".xlsx");
        dest.deleteOnExit();

        var convertor = new CsvToExcel(origin, dest.getAbsolutePath());
        convertor.setWorkbookType(WorkbookType.XSSF);
        convertor.convert();

        checkActualAndExpect(dest, "/example/result-all.xlsx", WorkbookType.XSSF);
    }

    @Test
    void sxssfTest() throws IOException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        var dest = File.createTempFile("dest-", ".xlsx");
        dest.deleteOnExit();

        var convertor = new CsvToExcel(origin, dest.getAbsolutePath());
        convertor.setWorkbookType(WorkbookType.SXSSF);
        convertor.convert();

        checkActualAndExpect(dest, "/example/result-all.xlsx", WorkbookType.SXSSF);
    }

    @Test
    void columnSelectTest() throws IOException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        var dest = File.createTempFile("dest-", ".xlsx");
        dest.deleteOnExit();

        var convertor = new CsvToExcel(origin, dest.getAbsolutePath());
        convertor.setWorkbookType(WorkbookType.SXSSF);
        convertor.convert(List.of(new Pair<>("Sheet1", List.of("번호", "이름"))));

        checkActualAndExpect(dest, "/example/result-num-name.xlsx", WorkbookType.SXSSF);
    }

    @Test
    void twoSheetColumnSelectTest() throws IOException, ClassNotFoundException, InvocationTargetException, NoSuchMethodException, InstantiationException, IllegalAccessException {
        var dest = File.createTempFile("dest-", ".xlsx");
        dest.deleteOnExit();

        var convertor = new CsvToExcel(origin, dest.getAbsolutePath());
        convertor.setWorkbookType(WorkbookType.SXSSF);
        convertor.convert(List.of(new Pair<>("Sheet1", List.of("번호", "이름")),
                new Pair<>("Sheet2", List.of("학년", "이름"))));

        checkActualAndExpect(dest, "/example/result-two-sheet.xlsx", WorkbookType.SXSSF);
    }
}