package com.yitiankeji.excel.utils;

import com.yitiankeji.excel.annotation.ExcelProperty;
import lombok.Data;

import java.lang.reflect.Field;
import java.util.*;

public class PropertyFieldSorter {

    public static List<Field> getIndexFields(Class<?> type) {
        List<Field> fields = getAllFields(type);
        return sortedFields(fields);
    }

    private static List<Field> getAllFields(Class<?> type) {
        List<Field> fields = new ArrayList<>(Arrays.asList(type.getDeclaredFields()));
        while ((type = type.getSuperclass()) != null) {
            fields.addAll(Arrays.asList(type.getDeclaredFields()));
        }
        return fields;
    }

    private static List<Field> sortedFields(List<Field> fields) {
        Map<Integer, Entry> indexFieldMap = new TreeMap<>();
        List<Field> unindexFields = new ArrayList<>(fields.size());

        int maxIndex = -1;
        for (Field field : fields) {
            ExcelProperty property = field.getAnnotation(ExcelProperty.class);
            if (property == null) {
                continue;
            }
            int index = property.index();
            if (index > maxIndex) {
                maxIndex = index;
            }

            if (index != -1) {
                Entry entry = indexFieldMap.get(index);
                if (entry == null) {
                    entry = new Entry(index);
                    indexFieldMap.put(index, entry);
                }
                entry.getFields().add(field);
            } else {
                unindexFields.add(field);
            }
        }

        Iterator<Field> unindexFieldIterator = unindexFields.iterator();
        List<Field> propertyFields = new ArrayList<>(fields.size());
        for (int i = 0; i <= Math.max(fields.size(), maxIndex); i++) {
            Entry entry = indexFieldMap.get(i);
            if (entry != null) {
                propertyFields.addAll(entry.getFields());
                continue;
            }
            boolean has = unindexFieldIterator.hasNext();
            if (!has) {
                continue;
            }
            propertyFields.add(unindexFieldIterator.next());
        }

        return propertyFields;
    }

    @Data
    static class Entry implements Comparable<Entry> {
        private int index;
        private List<Field> fields = new ArrayList<>();

        public Entry(int index) {
            this.index = index;
        }

        @Override
        public int compareTo(Entry o) {
            return index - o.index;
        }
    }
}
