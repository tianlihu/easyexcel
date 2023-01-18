package com.yitiankeji.excel.writer;

import org.apache.poi.xssf.usermodel.XSSFCell;

import java.lang.reflect.Field;

public interface ExcelWriteListener<T> {

    default boolean processHead(int rowIndex, int columnIndex, XSSFCell cell) {
        return false;
    }

    default boolean process(Object rowData, int rowIndex, int columnIndex, XSSFCell cell, Field field) {
        return false;
    }
}
