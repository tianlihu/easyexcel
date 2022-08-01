package com.yitiankeji.excel.utils;

import com.yitiankeji.excel.annotation.ExcelProperty;

import java.lang.reflect.Field;
import java.util.*;

public class PropertyFieldSorter {

    public static List<Field> getIndexFields(Class<?> type) {
        List<Field> fields = getAllFields(type);
        Map<Integer, Field> indexFieldMap = new HashMap<>();
        List<Field> unindexFields = new ArrayList<>(fields.size());
        List<Field> propertyFields = new ArrayList<>(fields.size());

        cacheFields(type, fields, indexFieldMap, unindexFields, propertyFields);

        List<Field> result = new ArrayList<>();
        int index = 0;
        int length = propertyFields.size();
        while (index < length) {
            Field field = indexFieldMap.get(index);
            if (field != null) {
                result.add(field);
                index++;
                continue;
            }

            field = unindexFields.get(0);
            result.add(field);
            unindexFields.remove(field);
            index++;
        }
        return result;
    }

    private static List<Field> getAllFields(Class<?> type) {
        List<Field> fields = new ArrayList<>(Arrays.asList(type.getDeclaredFields()));
        while ((type = type.getSuperclass()) != null) {
            fields.addAll(Arrays.asList(type.getDeclaredFields()));
        }
        return fields;
    }

    private static void cacheFields(Class<?> type, List<Field> fields, Map<Integer, Field> fieldMap, List<Field> unindexFields, List<Field> propertyFields) {
        for (Field field : fields) {
            ExcelProperty property = field.getAnnotation(ExcelProperty.class);
            if (property == null) {
                continue;
            }
            int index = property.index();
            if (index != -1) {
                Field cachedField = fieldMap.get(index);
                if (cachedField != null) {
                    throw new RuntimeException("Field property index ( " + index + " ) is already exists in Class " + type.getCanonicalName());
                }
            }

            if (index >= fields.size()) {
                throw new IndexOutOfBoundsException("Field property index ( " + index + " ) is out of bounds (" + fields.size() + ") in Class " + type.getCanonicalName());
            }
            fieldMap.put(index, field);
            if (index < 0) {
                unindexFields.add(field);
            }
            propertyFields.add(field);
        }
    }
}
