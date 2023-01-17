package com.yitiankeji.excel.converter;

import java.lang.reflect.Field;

public interface Converter<T> {

    Object convertToJavaData(String value, Field field);

    String convertToExcelData(T row, Object value, Field field);

    class AutoConverter<T> implements Converter<T> {

        @Override
        public String convertToExcelData(T row, Object value, Field field) {
            return null;
        }

        @Override
        public Object convertToJavaData(String value, Field field) {
            return null;
        }
    }
}
