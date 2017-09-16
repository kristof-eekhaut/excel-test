package be.eekhaut.kristof.exceltest;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.net.URL;
import java.util.List;

public abstract class ExcelTestCaseReader<T>  {

    private String fileName;

    public ExcelTestCaseReader(final String fileName) {
        super();
        this.fileName = fileName;
    }

    protected abstract List<T> createTestCases(Workbook workbook);

    public final List<T> createTestCases() throws IOException {

        InputStream inputStream = null;

        try {
            URL url = getFileUrl();
            inputStream = url.openStream();
            Workbook workbook = createWorkbook(inputStream);

            return createTestCases(workbook);
        } finally {
            if (inputStream != null) {
                inputStream.close();
            }
        }
    }

    private URL getFileUrl() {
        URL url = Thread.currentThread().getContextClassLoader().getResource(fileName);
        if (url == null) {
            throw new IllegalArgumentException("File not found: " + fileName);
        }
        return url;
    }

    private Workbook createWorkbook(InputStream inputStream) throws IOException {
        if(fileName.endsWith(".xls")) {
            return new HSSFWorkbook(inputStream);
        }
        if(fileName.endsWith(".xlsx")) {
            return new XSSFWorkbook(inputStream);
        }
        throw new IllegalStateException("Only .xls and .xlsx files are supported atm...");
    }
}

