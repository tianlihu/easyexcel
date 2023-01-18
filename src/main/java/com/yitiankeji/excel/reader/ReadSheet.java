package com.yitiankeji.excel.reader;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.utils.PropertyFieldSorter;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.experimental.Accessors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.*;

import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.math.BigDecimal;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;
import java.util.zip.DataFormatException;

@Data
@Accessors(fluent = true)
public class ReadSheet<T> {

    private int headRowNumber = 1;
    private Class<T> type;
    private Sheet sheet;
    private ExcelReadListener<T> listener;

    public List<T> doRead() {
        if (sheet == null) {
            return new ArrayList<>();
        }
        List<Field> fields = PropertyFieldSorter.getIndexFields(type);
        List<T> records = new ArrayList<>(1000);

        // 如果没有表头或没有内容，返回空列表
        int lastRowNum = sheet.getLastRowNum();
        if (headRowNumber >= lastRowNum) {
            return new ArrayList<>();
        }

        List<String> columnNames = readColumnNames();
        for (int rowIndex = headRowNumber; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            T rowData = readRow(row, fields, type, columnNames);
            if (listener != null) {
                listener.process(rowData, rowIndex);
            }
            records.add(rowData);
        }
        return records;
    }

    private List<String> readColumnNames() {
        List<String> columnNames = new ArrayList<>();
        Row headRow = sheet.getRow(headRowNumber - 1);
        for (int column = 0; column <= headRow.getLastCellNum(); column++) {
            String columnName = getCellValue(headRow.getCell(column));
            if (StringUtils.isNotEmpty(columnName)) {
                columnNames.add(StringUtils.trimToEmpty(columnName));
            }
        }
        return columnNames;
    }

    @SneakyThrows
    private T readRow(Row row, List<Field> fields, Class<T> type, List<String> columnNames) {
        Constructor<T> constructor = type.getConstructor();
        constructor.setAccessible(true);
        T instance = constructor.newInstance();
        for (Field field : fields) {
            int columnIndex = getColumnIndex(field, columnNames);
            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValue(cell);
            field.setAccessible(true);
            field.set(instance, convertValue(field, cellValue));
        }
        return instance;
    }

    private static int getColumnIndex(Field field, List<String> columnNames) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        String[] names = property.value();
        int columnIndex = property.index();
        for (String name : names) {
            int index = columnNames.indexOf(StringUtils.trimToNull(name));
            if (index == -1) {
                columnIndex = index;
            }
        }
        return columnIndex;
    }

    @SneakyThrows
    @SuppressWarnings("unchecked")
    private Object convertValue(Field field, String value) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        Class<? extends Converter> converter = property.converter();
        if (!converter.equals(Converter.AutoConverter.class)) {
            return convertValue(value, converter, field);
        }
        if (value == null) {
            return null;
        }

        Class<?> type = field.getType();
        if (type.equals(String.class)) {
            return value;
        }
        if (type.equals(Integer.class)) {
            String format = property.format();
            if (StringUtils.isNotEmpty(format)) {
                NumberFormat numberFormat = new DecimalFormat(format);
                return numberFormat.parse(value).intValue();
            }
            return Integer.valueOf(value);
        }
        if (type.equals(Long.class)) {
            String format = property.format();
            if (StringUtils.isNotEmpty(format)) {
                NumberFormat numberFormat = new DecimalFormat(format);
                return numberFormat.parse(value).intValue();
            }
            return Long.valueOf(value);
        }
        if (type.equals(BigDecimal.class)) {
            value = value.replaceAll(",", "");
            return new BigDecimal(value);
        }
        if (Date.class.isAssignableFrom(type)) {
            String format = property.format();
            if (StringUtils.isEmpty(format)) {
                throw new DataFormatException("日期格式需要定义");
            }
            DateFormat numberFormat = new SimpleDateFormat(format);
            return numberFormat.parse(value);
        }
        throw new RuntimeException("错误的值");
    }

    @SneakyThrows
    private Object convertValue(String value, Class<? extends Converter> converter, Field field) {
        Constructor<? extends Converter> constructor = converter.getDeclaredConstructor();
        constructor.setAccessible(true);
        Converter convertor = constructor.newInstance();
        return convertor.convertToJavaData(value, field);
    }

    public static String getCellValue(Cell cell) {
        //判断是否为null或空串
        if (cell == null || cell.toString().trim().equals("")) {
            return "";
        }
        CellType cellType = cell.getCellTypeEnum();
        switch (cellType) {
            case NUMERIC: // 数字
                short format = cell.getCellStyle().getDataFormat();
                if (DateUtil.isCellDateFormatted(cell)) { // 日期
                    SimpleDateFormat sdf;
                    if (format == 20 || format == 32) {
                        sdf = new SimpleDateFormat("HH:mm");
                    } else if (format == 14 || format == 31 || format == 57 || format == 58) {
                        sdf = new SimpleDateFormat("yyyy-MM-dd");
                    } else if (format == 179) {
                        sdf = new SimpleDateFormat("HH:mm:ss");
                    } else {
                        sdf = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
                    }
                    return sdf.format(cell.getDateCellValue());
                } else {
                    cell.setCellType(CellType.STRING);
                    return cell.getStringCellValue();
                }
            case STRING: // 字符串
                return cell.getStringCellValue();
            case BOOLEAN: // Boolean
                return cell.getBooleanCellValue() + "";
            case FORMULA: // 公式
                cell.setCellType(CellType.STRING);
                return cell.getStringCellValue();
            case BLANK: // 空值
                return "";
            case ERROR: // 故障
                return "ERROR VALUE";
            default:
                return "UNKNOWN VALUE";
        }
    }
}
