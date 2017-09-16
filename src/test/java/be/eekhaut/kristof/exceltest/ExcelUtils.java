package be.eekhaut.kristof.exceltest;

import org.apache.commons.lang3.BooleanUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.RichTextString;
import org.apache.poi.ss.usermodel.Row;

import java.math.BigDecimal;

import static org.apache.commons.lang3.StringUtils.isBlank;

public final class ExcelUtils {

    private ExcelUtils(){

    }

    public static String getString(final Cell cell) {
        if (null != cell && cell.getCellTypeEnum() != CellType.BLANK) {
            if (cell.getCellTypeEnum() == CellType.STRING || cell.getCellTypeEnum() == CellType.FORMULA) {
                RichTextString value = cell.getRichStringCellValue();
                if (null != value) {
                    return value.getString();
                }
            } else {
                Double value = cell.getNumericCellValue();
                return "" + value.longValue();
            }
        }
        return null;
    }

    public static <E extends Enum<E>> E getEnum(final Cell cell, final Class<E> enumClass) {
        if (cell == null || cell.getRichStringCellValue() == null) {
            return null;
        }
        String constantName = cell.getRichStringCellValue().getString();
        if (isBlank(constantName)) {
            return null;
        }
        return Enum.valueOf(enumClass, constantName);
    }

    public static boolean getBoolean(final Cell cell) {
        String value = getString(cell);
        if (isBlank(value)) {
            return false;
        }
        return value.equalsIgnoreCase("Y") || BooleanUtils.toBoolean(value);
    }

    public static Double getDouble(final Cell cell) {
        if (null != cell && cell.getCellTypeEnum() != CellType.BLANK) {
            if (cell.getCellTypeEnum() == CellType.STRING) {
                RichTextString value = cell.getRichStringCellValue();
                if (null != value) {
                    return Double.parseDouble(value.getString());
                }
            } else {
                return cell.getNumericCellValue();
            }
        }
        return null;
    }

    public static Long getLong(final Cell cell) {
        Double value = getDouble(cell);
        return (value == null) ? null : value.longValue();
    }

    public static Integer getInteger(final Cell cell) {
        Double value = getDouble(cell);
        return (value == null) ? null : value.intValue();
    }

    public static BigDecimal getBigDecimal(final Cell cell) {
        String string = getString(cell);

        return (isBlank(string)) ? null : new BigDecimal(string);
    }

    public static int getRowNum(final Row row) {
        return row.getRowNum() + 1;
    }
}

