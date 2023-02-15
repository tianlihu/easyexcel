package com.yitiankeji.excel.writer;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.utils.PropertyFieldSorter;
import lombok.SneakyThrows;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.*;

import java.io.BufferedOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.lang.reflect.Constructor;
import java.lang.reflect.Field;
import java.text.DateFormat;
import java.text.DecimalFormat;
import java.text.NumberFormat;
import java.text.SimpleDateFormat;
import java.util.*;

import static org.apache.poi.ss.usermodel.BorderStyle.THIN;
import static org.apache.poi.ss.usermodel.IndexedColors.BLACK1;

public class ExcelWriter {

    private final OutputStream output;
    private final List<WriteSheet> writeSheets = new ArrayList<>();

    public ExcelWriter(OutputStream output) {
        this.output = new BufferedOutputStream(output);
    }

    public ExcelWriter sheet(Class<?> type, List<?> records) {
        return sheet(null, type, records);
    }

    public ExcelWriter sheet(String sheetName, Class<?> type, List<?> records) {
        WriteSheet writeSheet = new WriteSheet();
        writeSheet.setSheetName(sheetName);
        writeSheet.setType(type);
        writeSheet.setRecords(records);
        writeSheets.add(writeSheet);
        return this;
    }

    public void doWrite(Class<?> type, List<?> records) throws IOException {
        sheet(type, records);
        doWrite();
    }

    public void doWrite(String sheetName, Class<?> type, List<?> records) throws IOException {
        sheet(sheetName, type, records);
        doWrite();
    }

    public void doWrite() throws IOException {
        doWrite(null);
    }

    public void doWrite(ExcelWriteListener<?> listener) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(XSSFWorkbookType.XLSX)) {
            XSSFCellStyle titleStyle = createTitleStyle(workbook);
            XSSFCellStyle cellStyle = createCellStyle(workbook);

            // 逐个Sheet写入文档
            for (int i = 0; i < writeSheets.size(); i++) {
                WriteSheet writeSheet = writeSheets.get(i);
                XSSFSheet sheet = workbook.createSheet(writeSheet.getSheetName() != null ? writeSheet.getSheetName() : "Sheet" + (i + 1));

                // 写入表头
                List<Field> fields = PropertyFieldSorter.getIndexFields(writeSheet.getType());
                Map<String, XSSFCellStyle> cellStyleMap = new HashMap<>(fields.size());
                XSSFRow headRow = sheet.createRow(0);
                for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                    Field field = fields.get(columnIndex);
                    ExcelProperty property = field.getAnnotation(ExcelProperty.class);
                    XSSFCell cell = headRow.createCell(columnIndex);
                    cell.setCellStyle(titleStyle);
                    cellStyleMap.put(field.getName(), createCellStyle(workbook, field, cellStyle));
                    if (listener != null) {
                        boolean process = listener.processHead(0, columnIndex, cell);
                        if (!process) {
                            cell.setCellValue(property.value()[0]);
                        }
                    } else {
                        cell.setCellValue(property.value()[0]);
                    }
                }

                // 逐行写入当前Sheet
                List<?> records = writeSheet.getRecords();
                for (int rowIndex = 0; rowIndex < records.size(); rowIndex++) {
                    XSSFRow row = sheet.createRow(rowIndex + 1);
                    Object record = records.get(rowIndex);
                    writeRow(workbook, sheet, record, fields, row, rowIndex + 1, listener, cellStyleMap);
                }

                // 列宽自适应
                for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                    sheet.autoSizeColumn(columnIndex, true);
                }
            }

            // 保存文档
            workbook.write(output);
            output.flush();
        } finally {
            IOUtils.closeQuietly(output);
        }
    }

    private static XSSFCellStyle createTitleStyle(XSSFWorkbook workbook) {
        XSSFCellStyle titleStyle = workbook.createCellStyle();
        titleStyle.setBorderLeft(THIN);
        titleStyle.setBorderRight(THIN);
        titleStyle.setBorderTop(THIN);
        titleStyle.setBorderBottom(THIN);
        titleStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        titleStyle.setFillForegroundColor((short) 22);
        titleStyle.setAlignment(HorizontalAlignment.CENTER);

        XSSFFont titleFont = workbook.createFont();
        titleFont.setColor(BLACK1.index);
        titleFont.setBold(true);
        titleStyle.setFont(titleFont);
        return titleStyle;
    }

    private XSSFCellStyle createCellStyle(XSSFWorkbook workbook, Field field, XSSFCellStyle defaultCellStyle) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        if (StringUtils.isEmpty(property.format())) {
            return defaultCellStyle;
        }

        XSSFCellStyle cellStyle = createCellStyle(workbook);
        XSSFDataFormat dataFormat = workbook.createDataFormat();
        cellStyle.setDataFormat(dataFormat.getFormat(property.format()));
        return cellStyle;
    }

    private static XSSFCellStyle createCellStyle(XSSFWorkbook workbook) {
        XSSFCellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setBorderLeft(THIN);
        cellStyle.setBorderRight(THIN);
        cellStyle.setBorderTop(THIN);
        cellStyle.setBorderBottom(THIN);
        return cellStyle;
    }

    @SneakyThrows
    private void writeRow(XSSFWorkbook workbook, XSSFSheet sheet, Object record, List<Field> fields, XSSFRow row, int rowIndex, ExcelWriteListener<?> listener, Map<String, XSSFCellStyle> cellStyleMap) {
        for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
            XSSFCell cell = row.createCell(columnIndex);
            Field field = fields.get(columnIndex);
            cell.setCellStyle(cellStyleMap.get(field.getName()));
            field.setAccessible(true);
            Object value = field.get(record);
            if (listener == null) {
                writeCell(cell, field, record, value);
                continue;
            }

            boolean process = listener.process(record, rowIndex, columnIndex, workbook, sheet, cell, field);
            if (process) {
                continue;
            }

            writeCell(cell, field, record, value);
        }
    }

    @SneakyThrows
    private void writeCell(XSSFCell cell, Field field, Object record, Object value) {
        if (value == null) {
            return;
        }

        Class<?> fieldType = field.getType();
        if (boolean.class.equals(fieldType) || Boolean.class.equals(fieldType)) {
            if (value instanceof Boolean) {
                cell.setCellValue((Boolean) value);
            } else if (value instanceof String) {
                cell.setCellValue(Boolean.parseBoolean((String) value));
            } else {
                throw new RuntimeException("布尔类型列，只能是true或false：" + field);
            }
            return;
        }

        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        if (StringUtils.isEmpty(property.format())) {
            if (value instanceof Number) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (value instanceof Date) {
                cell.setCellValue(((Date) value));
            } else {
                cell.setCellValue(getFieldValue(record, field));
            }
            return;
        }

        if (property.type() == ExcelProperty.NUMBER) {
            if (Number.class.isAssignableFrom(fieldType)) {
                cell.setCellValue(((Number) value).doubleValue());
            } else if (String.class.equals(fieldType)) {
                DecimalFormat format = new DecimalFormat(property.format());
                Number number = format.parse((String) value);
                cell.setCellValue(number.doubleValue());
            } else {
                throw new RuntimeException("数值类型列，只能是数字类型或字符串类型：" + field);
            }
        } else if (property.type() == ExcelProperty.DATE) {
            if (Date.class.isAssignableFrom(fieldType)) {
                DateFormat format = new SimpleDateFormat(property.format());
                cell.setCellValue(format.parse(format.format(value)));
            } else if (String.class.equals(fieldType)) {
                DateFormat format = new SimpleDateFormat(property.format());
                Date date = format.parse((String) value);
                cell.setCellValue(date);
            } else {
                throw new RuntimeException("日期类型列，只能是日期类型或字符串类型：" + field);
            }
        } else {
            cell.setCellValue(getFieldValue(record, field));
        }
    }

    @SneakyThrows
    private String getFieldValue(Object record, Field field) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        field.setAccessible(true);
        Object value = field.get(record);
        Class<? extends Converter> converter = property.converter();
        if (!converter.equals(Converter.AutoConverter.AutoConverter.class)) {
            return convertValue(record, value, converter, field);
        }
        if (value == null) {
            return null;
        }

        if (value instanceof String) {
            return (String) value;
        }
        if (value instanceof Number) {
            String format = property.format();
            if (StringUtils.isNotEmpty(format)) {
                NumberFormat numberFormat = new DecimalFormat(format);
                return numberFormat.format(value);
            }
            return value.toString();
        }
        if (value instanceof Date) {
            String format = property.format();
            if (StringUtils.isNotEmpty(format)) {
                DateFormat numberFormat = new SimpleDateFormat(format);
                return numberFormat.format(value);
            }
        }
        return value.toString();
    }

    @SneakyThrows
    private String convertValue(Object row, Object value, Class<? extends Converter> converter, Field field) {
        Constructor<? extends Converter> constructor = converter.getDeclaredConstructor();
        constructor.setAccessible(true);
        Converter convertor = constructor.newInstance();
        return convertor.convertToExcelData(row, value, field);
    }
}
