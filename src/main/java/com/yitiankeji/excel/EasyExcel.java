package com.yitiankeji.excel;

import com.yitiankeji.excel.annotation.ExcelProperty;
import com.yitiankeji.excel.reader.ExcelReader;
import com.yitiankeji.excel.writer.ExcelWriter;
import lombok.Data;
import lombok.SneakyThrows;

import java.io.*;
import java.nio.file.Files;
import java.util.ArrayList;
import java.util.List;

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

    @SneakyThrows
    public static void main(String[] args) {
        List<User> users = new ArrayList<>();
        for (int i = 0; i < 100; i++) {
            User user = new User();
            user.setId(i);
            user.setName("name_" + i);
            user.setAge(i * 10);
            users.add(user);
        }
        EasyExcel.write("C:/Users/tianlihu/Desktop/test.xlsx").doWrite(User.class, users);

//        List<User> users = EasyExcel.read("C:/Users/tianlihu/Desktop/test.xlsx").doReadAll(User.class);
//        users.forEach(System.out::println);
    }
}

@Data
class User {
    private Integer id;
    @ExcelProperty("姓名")
    private String name;
    @ExcelProperty("年龄")
    private Integer age;
}