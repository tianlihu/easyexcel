package com.yitiankeji.excel.writer;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.utils.PropertyFieldSorter;
import lombok.SneakyThrows;
import org.apache.commons.io.IOUtils;
import org.apache.commons.lang3.StringUtils;
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
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

public class ExcelWriter<T> {

    private final OutputStream output;
    private final List<WriteSheet<T>> writeSheets = new ArrayList<>();

    public ExcelWriter(OutputStream output) {
        this.output = new BufferedOutputStream(output);
    }

    public ExcelWriter<T> sheet(Class<T> type, List<T> records) {
        return sheet(null, type, records);
    }

    public ExcelWriter<T> sheet(String sheetName, Class<T> type, List<T> records) {
        WriteSheet<T> writeSheet = new WriteSheet<>();
        writeSheet.setSheetName(sheetName);
        writeSheet.setType(type);
        writeSheet.setRecords(records);
        writeSheets.add(writeSheet);
        return this;
    }

    public void doWrite(Class<T> type, List<T> records) throws IOException {
        sheet(type, records);
        doWrite();
    }

    public void doWrite(String sheetName, Class<T> type, List<T> records) throws IOException {
        sheet(sheetName, type, records);
        doWrite();
    }

    public void doWrite() throws IOException {
        doWrite(null);
    }

    public void doWrite(ExcelWriteListener<T> listener) throws IOException {
        try (XSSFWorkbook workbook = new XSSFWorkbook(XSSFWorkbookType.XLSX)) {
            // 逐个Sheet写入文档
            for (int i = 0; i < writeSheets.size(); i++) {
                WriteSheet<T> writeSheet = writeSheets.get(i);
                XSSFSheet sheet = workbook.createSheet(writeSheet.getSheetName() != null ? writeSheet.getSheetName() : "Sheet" + (i + 1));

                // 写入表头
                List<Field> fields = PropertyFieldSorter.getIndexFields(writeSheet.getType());
                XSSFRow headRow = sheet.createRow(0);
                for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                    Field field = fields.get(columnIndex);
                    ExcelProperty property = field.getAnnotation(ExcelProperty.class);
                    XSSFCell cell = headRow.createCell(columnIndex);
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
                List<T> records = writeSheet.getRecords();
                for (int rowIndex = 0; rowIndex < records.size(); rowIndex++) {
                    XSSFRow row = sheet.createRow(rowIndex + 1);
                    T record = records.get(rowIndex);
                    writeRow(record, fields, row, rowIndex + 1, listener);
                }

                // 列宽自适应
                for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
                    sheet.autoSizeColumn(columnIndex);
                }
            }

            // 保存文档
            workbook.write(output);
            output.flush();
        } finally {
            IOUtils.closeQuietly(output);
        }
    }

    private void writeRow(T record, List<Field> fields, XSSFRow row, int rowIndex, ExcelWriteListener<T> listener) {
        for (int columnIndex = 0; columnIndex < fields.size(); columnIndex++) {
            XSSFCell cell = row.createCell(columnIndex);
            Field field = fields.get(columnIndex);
            if (listener == null) {
                cell.setCellValue(getFieldValue(record, field));
                return;
            }

            boolean process = listener.process(record, rowIndex, columnIndex, cell, field);
            if (process) {
                return;
            }

            cell.setCellValue(getFieldValue(record, field));
        }
    }

    @SuppressWarnings("unchecked")
    @SneakyThrows
    private String getFieldValue(T record, Field field) {
        ExcelProperty property = field.getAnnotation(ExcelProperty.class);
        field.setAccessible(true);
        Object value = field.get(record);
        Class<? extends Converter<T>> converter = (Class<? extends Converter<T>>) property.converter();
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
    private String convertValue(T row, Object value, Class<? extends Converter<T>> converter, Field field) {
        Constructor<? extends Converter<T>> constructor = converter.getDeclaredConstructor();
        constructor.setAccessible(true);
        Converter<T> convertor = constructor.newInstance();
        return convertor.convertToExcelData(row, value, field);
    }
}
