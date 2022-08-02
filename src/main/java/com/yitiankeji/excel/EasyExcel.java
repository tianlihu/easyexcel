package com.yitiankeji.excel;

import com.yitiankeji.excel.reader.ExcelReader;
import com.yitiankeji.excel.writer.ExcelWriter;

import java.io.*;
import java.nio.file.Files;

public class EasyExcel {

    public static ExcelReader read(InputStream input) throws IOException {
        return new ExcelReader(input instanceof BufferedInputStream ? (BufferedInputStream) input : new BufferedInputStream(input));
    }

    public static ExcelReader read(String filePath) throws IOException {
        return read(new File(filePath));
    }

    public static <T> ExcelReader read(File file) throws IOException {
        return read(new FileInputStream(file));
    }

    public static ExcelWriter write(OutputStream output) {
        return new ExcelWriter(output);
    }

    public static ExcelWriter write(String filePath) throws IOException {
        File file = new File(filePath);
        file.getParentFile().mkdirs();
        return write(Files.newOutputStream(file.toPath()));
    }

    public static ExcelWriter write(File file) throws IOException {
        file.getParentFile().mkdirs();
        return write(Files.newOutputStream(file.toPath()));
    }
}