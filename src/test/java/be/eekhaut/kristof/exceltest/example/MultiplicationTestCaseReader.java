package be.eekhaut.kristof.exceltest.example;

import be.eekhaut.kristof.exceltest.ExcelTestCaseReader;
import lombok.AllArgsConstructor;
import lombok.Getter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.ArrayList;
import java.util.Iterator;
import java.util.List;

import static be.eekhaut.kristof.exceltest.ExcelUtils.getBoolean;
import static be.eekhaut.kristof.exceltest.ExcelUtils.getInteger;

public class MultiplicationTestCaseReader extends ExcelTestCaseReader<MultiplicationTestCaseReader.TestCase> {

    private static int column = 0;

    // Test Case
    private static final int COL_IS_ENABLED = column++;
    private static final int COL_VALUE = column++;
    private static final int COL_MULTIPLIER = column++;
    private static final int COL_RESULT = column++;

    MultiplicationTestCaseReader(String fileName) {
        super(fileName);
    }

    @Override
    protected List<TestCase> createTestCases(Workbook workbook) {
        List<TestCase> testCases = new ArrayList<>();
        Iterator<Row> rowIterator = workbook.getSheetAt(0).iterator();
        rowIterator.next();// Skip header row

        while (rowIterator.hasNext()) {

            Row row = rowIterator.next();

            if(getBoolean(row.getCell(COL_IS_ENABLED))) {
                int value = getInteger(row.getCell(COL_VALUE));
                int multiplier = getInteger(row.getCell(COL_MULTIPLIER));
                int result = getInteger(row.getCell(COL_RESULT));

                testCases.add(new TestCase(value, multiplier, result));
            }
        }

        return testCases;
    }

    @Getter
    @AllArgsConstructor
    static class TestCase {

        private int value;
        private int multiplier;
        private int result;
    }
}
