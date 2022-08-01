package com.yitiankeji.excel.reader;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.utils.PropertyFieldSorter;
import lombok.Data;
import lombok.SneakyThrows;
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
public class ReadSheet<T> {

    private int headRowNumber = 1;
    private Class<T> type;
    private Sheet sheet;
    private ExcelReadListener<T> listener;

    public List<T> doRead() {
        List<Field> fields = PropertyFieldSorter.getIndexFields(type);
        List<T> records = new ArrayList<>(1000);
        int lastRowNum = sheet.getLastRowNum();
        for (int rowIndex = headRowNumber; rowIndex < lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            records.add(readRow(row, fields, type));
        }
        return records;
    }

    @SneakyThrows
    private T readRow(Row row, List<Field> fields, Class<T> type) {
        Constructor<T> constructor = type.getConstructor();
        constructor.setAccessible(true);
        T instance = constructor.newInstance();
        for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
            Cell cell = row.getCell(columnIndex);
            String cellValue = getCellValue(cell);
            Field field = fields.get(columnIndex);
            field.setAccessible(true);
            field.set(instance, convertValue(field, cellValue));
        }
        return instance;
    }

    @SneakyThrows
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

    @SuppressWarnings("deprecation")
    public static String getCellValue(Cell cell) {
        //判断是否为null或空串
        if (cell == null || cell.toString().trim().equals("")) {
            return "";
        }
        CellType cellType = cell.getCellType();
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
