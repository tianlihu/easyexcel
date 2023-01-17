package com.yitiankeji.excel;

import com.yitiankeji.excel.reader.ExcelReader;
import com.yitiankeji.excel.writer.ExcelWriter;

import java.io.*;
import java.nio.file.Files;

public class EasyExcel {

    public static <T> ExcelReader<T> read(InputStream input) throws IOException {
        return new ExcelReader<>(input instanceof BufferedInputStream ? (BufferedInputStream) input : new BufferedInputStream(input));
    }

    public static <T> ExcelReader<T> read(String filePath) throws IOException {
        return read(new File(filePath));
    }

    public static <T> ExcelReader<T> read(File file) throws IOException {
        return read(Files.newInputStream(file.toPath()));
    }

    public static <T> ExcelWriter<T> write(OutputStream output) {
        return new ExcelWriter<T>(output);
    }

    public static <T> ExcelWriter<T> write(String filePath) throws IOException {
        File file = new File(filePath);
        file.getParentFile().mkdirs();
        return write(Files.newOutputStream(file.toPath()));
    }

    public static <T> ExcelWriter<T> write(File file) throws IOException {
        file.getParentFile().mkdirs();
        return write(Files.newOutputStream(file.toPath()));
    }
}