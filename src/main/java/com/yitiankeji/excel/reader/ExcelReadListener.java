package com.yitiankeji.excel.reader;

import java.util.List;

public interface ExcelReadListener<T> {

    default void processHead(List<String> heads, int rowIndex) {
    }

    default void process(T rowData, int rowIndex) {
    }
}
