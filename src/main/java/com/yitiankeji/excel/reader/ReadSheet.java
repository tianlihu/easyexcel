package com.yitiankeji.excel.reader;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.utils.PropertyFieldSorter;
import lombok.Data;
import lombok.SneakyThrows;
import lombok.experimental.Accessors;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;

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
        if (headRowNumber > lastRowNum) {
            return new ArrayList<>();
        }

        List<String> columnNames = readColumnNames();
        for (int rowIndex = headRowNumber; rowIndex <= lastRowNum; rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null) { // 跳过空行
                continue;
            }
            T rowData = readRow(row, fields, type, columnNames);
            if (rowData == null) { // 跳过空行
                continue;
            }
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
            Cell cell = headRow.getCell(column);
            String columnName = cell == null ? "" : cell.getStringCellValue();
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
        boolean emptyLine = true;
        for (Field field : fields) {
            field.setAccessible(true);

            int columnIndex = getColumnIndex(field, columnNames);
            Cell cell = row.getCell(columnIndex);
            if (cell == null) {
                continue;
            }
            Object cellValue = getCellValue(field, cell);
            field.set(instance, convertValue(field, cellValue));

            String strValue = cellValue == null ? "" : cellValue.toString();  // 跳过空行
            if (!StringUtils.isEmpty(strValue)) {
                emptyLine = false;
            }
        }

        if (emptyLine) {
            return null;
        }

        return instance;
    }

    private static int getColumnIndex(Field field, List<String> columnNames) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        String[] names = property.value();
        int columnIndex = property.index();
        for (String name : names) {
            int index = columnNames.indexOf(StringUtils.trimToNull(name));
            if (index != -1) {
                columnIndex = index;
            }
        }
        return columnIndex;
    }

    @SneakyThrows
    private Object convertValue(Field field, Object value) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        Class<? extends Converter> converter = property.converter();
        if (!converter.equals(Converter.AutoConverter.class)) {
            return convertValue(value, converter, field);
        }
        if (value == null) {
            return null;
        }

        Class<?> fieldType = field.getType();
        String format = property.format();
        if (StringUtils.isEmpty(format)) {
            if (String.class.equals(fieldType)) {
                return value.toString();
            } else if (boolean.class.equals(fieldType) || Boolean.class.equals(fieldType)) {
                if (value instanceof Boolean) {
                    return value;
                } else if (value instanceof String) {
                    return Boolean.valueOf((String) value);
                }
            } else if (BigDecimal.class.equals(fieldType)) {
                if (value instanceof Number) {
                    return BigDecimal.valueOf(((Number) value).doubleValue());
                } else if (value instanceof String) {
                    return new BigDecimal((String) value);
                }
            } else if (short.class.equals(fieldType) || Short.class.equals(fieldType)) {
                NumberFormat numberFormat = new DecimalFormat("#.00");
                if (value instanceof Number) {
                    return ((Number) value).intValue();
                } else if (value instanceof String) {
                    return numberFormat.parse((String) value).intValue();
                }
            } else if (int.class.equals(fieldType) || Integer.class.equals(fieldType)) {
                NumberFormat numberFormat = new DecimalFormat("#.00");
                if (value instanceof Number) {
                    return ((Number) value).intValue();
                } else if (value instanceof String) {
                    return numberFormat.parse((String) value).intValue();
                }
            } else if (long.class.equals(fieldType) || Long.class.equals(fieldType)) {
                NumberFormat numberFormat = new DecimalFormat("#.00");
                if (value instanceof Number) {
                    return ((Number) value).longValue();
                } else if (value instanceof String) {
                    return numberFormat.parse((String) value).longValue();
                }
            } else if (float.class.equals(fieldType) || Float.class.equals(fieldType)) {
                NumberFormat numberFormat = new DecimalFormat("#.00");
                if (value instanceof Number) {
                    return ((Number) value).floatValue();
                } else if (value instanceof String) {
                    return numberFormat.parse((String) value).floatValue();
                }
            } else if (double.class.equals(fieldType) || Double.class.equals(fieldType)) {
                NumberFormat numberFormat = new DecimalFormat("#.00");
                if (value instanceof Number) {
                    return ((Number) value).doubleValue();
                } else if (value instanceof String) {
                    return numberFormat.parse((String) value).doubleValue();
                }
            } else if (Date.class.equals(fieldType)) {
                if (value instanceof Date) {
                    return value;
                } else if (value instanceof String) {
                    String date = (String) value;
                    date = date.replaceAll("-", "").replaceAll(":", "").replaceAll(" ", "");
                    if (date.length() == 8) {
                        DateFormat dateFormat = new SimpleDateFormat("yyyyMMdd");
                        return dateFormat.parse(date);
                    } else if (date.length() == 14) {
                        DateFormat dateFormat = new SimpleDateFormat("yyyyMMddHHmmss");
                        return dateFormat.parse(date);
                    }
                }
            }
            return value;
        }

        if (String.class.equals(fieldType)) {
            if (property.type() == ExcelProperty.NUMBER) {
                if (value instanceof String) {
                    return value;
                } else if (value instanceof Number) {
                    NumberFormat numberFormat = new DecimalFormat(format);
                    return numberFormat.format(value);
                }
            } else if (property.type() == ExcelProperty.DATE) {
                if (value instanceof String) {
                    return value;
                } else if (value instanceof Date) {
                    DateFormat dateFormat = new SimpleDateFormat(format);
                    return dateFormat.format(value);
                }
            }
            return value;
        } else if (BigDecimal.class.equals(fieldType)) {
            if (value instanceof Number) {
                return BigDecimal.valueOf(((Number) value).doubleValue());
            } else if (value instanceof String) {
                return new BigDecimal((String) value);
            }
        } else if (short.class.equals(fieldType) || Short.class.equals(fieldType)) {
            NumberFormat numberFormat = new DecimalFormat(format);
            if (value instanceof Number) {
                return ((Number) value).intValue();
            } else if (value instanceof String) {
                return numberFormat.parse((String) value).intValue();
            }
        } else if (int.class.equals(fieldType) || Integer.class.equals(fieldType)) {
            NumberFormat numberFormat = new DecimalFormat(format);
            if (value instanceof Number) {
                return ((Number) value).intValue();
            } else if (value instanceof String) {
                return numberFormat.parse((String) value).intValue();
            }
        } else if (long.class.equals(fieldType) || Long.class.equals(fieldType)) {
            NumberFormat numberFormat = new DecimalFormat(format);
            if (value instanceof Number) {
                return ((Number) value).longValue();
            } else if (value instanceof String) {
                return numberFormat.parse((String) value).longValue();
            }
        } else if (float.class.equals(fieldType) || Float.class.equals(fieldType)) {
            NumberFormat numberFormat = new DecimalFormat(format);
            if (value instanceof Number) {
                return ((Number) value).floatValue();
            } else if (value instanceof String) {
                return numberFormat.parse((String) value).floatValue();
            }
        } else if (double.class.equals(fieldType) || Double.class.equals(fieldType)) {
            NumberFormat numberFormat = new DecimalFormat(format);
            if (value instanceof Number) {
                return ((Number) value).doubleValue();
            } else if (value instanceof String) {
                return numberFormat.parse((String) value).doubleValue();
            }
        } else if (Date.class.equals(fieldType)) {
            DateFormat dateFormat = new SimpleDateFormat(format);
            if (value instanceof Date) {
                return value;
            } else if (value instanceof String) {
                return dateFormat.parse((String) value);
            }
        } else if (boolean.class.equals(fieldType) || Boolean.class.equals(fieldType)) {
            if (value instanceof Boolean) {
                return value;
            } else if (value instanceof String) {
                return Boolean.valueOf((String) value);
            }
        }
        throw new RuntimeException("错误的值:" + value + ", " + field);
    }

    @SneakyThrows
    private Object convertValue(Object value, Class<? extends Converter> converter, Field field) {
        Constructor<? extends Converter> constructor = converter.getDeclaredConstructor();
        constructor.setAccessible(true);
        Converter convertor = constructor.newInstance();
        return convertor.convertToJavaData(value, field);
    }

    public static Object getCellValue(Field field, Cell cell) {
        //判断是否为null或空串
        if (cell == null || cell.toString().trim().equals("")) {
            return null;
        }
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        CellType cellType = cell.getCellTypeEnum();
        Class<?> fieldType = field.getType();
        switch (cellType) {
            case NUMERIC: // 数字
                if (Date.class.equals(fieldType) || property.type() == ExcelProperty.DATE) {
                    return cell.getDateCellValue();
                }

                double numericCellValue = cell.getNumericCellValue();
                String value = String.valueOf(numericCellValue);
                if (value.endsWith(".0")) {
                    return (long) numericCellValue;
                }
                return numericCellValue;
            case STRING: // 字符串
                return cell.getStringCellValue();
            case BOOLEAN: // Boolean
                return cell.getBooleanCellValue();
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
