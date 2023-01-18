package com.yitiankeji.excel.converter;

import java.lang.reflect.Field;

public interface Converter {

    Object convertToJavaData(String value, Field field);

    String convertToExcelData(Object row, Object value, Field field);

    class AutoConverter implements Converter {

        @Override
        public String convertToExcelData(Object row, Object value, Field field) {
            return null;
        }

        @Override
        public Object convertToJavaData(String value, Field field) {
            return null;
        }
    }
}
