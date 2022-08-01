package com.yitiankeji.excel.writer;

import lombok.Data;

import java.util.ArrayList;
import java.util.List;

@Data
public class WriteSheet {

    private String sheetName;
    private Class<?> type;
    private List<?> records = new ArrayList<>();
}
