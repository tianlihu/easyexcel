package com.yitiankeji.excel.utils;

import com.yitiankeji.excel.constants.Constants;
import lombok.SneakyThrows;
import org.apache.poi.poifs.filesystem.FileMagic;

import java.io.BufferedInputStream;

public class ExcelType {

    @SneakyThrows
    public static String type(BufferedInputStream input) {
        FileMagic fileMagic = FileMagic.valueOf(input);
        if (FileMagic.OLE2.equals(fileMagic)) {
            return Constants.XLS;
        } else if (FileMagic.OOXML.equals(fileMagic)) {
            return Constants.XLSX;
        }
        throw new RuntimeException("错误的Excel格式，请手工指定");
    }
}
