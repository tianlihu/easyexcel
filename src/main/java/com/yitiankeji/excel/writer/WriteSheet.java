package com.yitiankeji.excel.writer;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class WriteSheet<T> {

    private String sheetName;
    private Class<T> type;
    private List<T> records = new ArrayList<>();
}
