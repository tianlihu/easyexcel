package com.yitiankeji.excel.reader;

import java.util.Map;

public interface ExcelReadListener<T> {

    void processHead(Map<Integer, Object> headMap, int rowIndex);

    void process(T rowData, int rowIndex);
}
