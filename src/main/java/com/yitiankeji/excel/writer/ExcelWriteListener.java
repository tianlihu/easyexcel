package com.yitiankeji.excel.writer;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.lang.reflect.Field;

public interface ExcelWriteListener<T> {

    default boolean processHead(int rowIndex, int columnIndex, XSSFCell cell) {
        return false;
    }

    default boolean process(Object rowData, int rowIndex, int columnIndex, XSSFWorkbook workbook, XSSFSheet sheet, XSSFCell cell, Field field) {
        return false;
    }
}
