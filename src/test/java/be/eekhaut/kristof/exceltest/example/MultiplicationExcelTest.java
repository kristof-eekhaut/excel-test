package be.eekhaut.kristof.exceltest.example;

import org.junit.Test;
import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;

import java.io.IOException;
import java.util.List;

import static org.hamcrest.core.IsEqual.equalTo;
import static org.junit.Assert.assertThat;

@RunWith(Parameterized.class)
public class MultiplicationExcelTest {

    private MultiplicationTestCaseReader.TestCase testCase;

    public MultiplicationExcelTest(MultiplicationTestCaseReader.TestCase testCase) {
        this.testCase = testCase;
    }

    @Parameterized.Parameters
    public static List<MultiplicationTestCaseReader.TestCase> loadDataFromExcelFile() throws IOException {
        return new MultiplicationTestCaseReader("be/eekhaut/kristof/exceltest/example/MultiplicationTestCases.xlsx").createTestCases();
    }

    @Test
    public void testMultiplication() {
        assertThat(testCase.getValue() * testCase.getMultiplier(), equalTo(testCase.getResult()));
    }
}
