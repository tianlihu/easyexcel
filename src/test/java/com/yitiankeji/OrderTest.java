package com.yitiankeji;

import com.yitiankeji.excel.EasyExcel;
import lombok.SneakyThrows;

import java.math.BigDecimal;
import java.util.*;

public class OrderTest {

    @SneakyThrows
    public static void main(String[] args) {
        Order order = new Order();
        order.setField1("南方广利回报债券型证券投资基金A级222");
        order.setField2(new Date());
        order.setField3("20230217");
        order.setField4(10);
        order.setField5(3618287.74);
        order.setField6(new BigDecimal("3618287.74"));
        order.setField7("3618287.74");
        order.setField8(true);

        List<Order> orders = new ArrayList<>();
        orders.add(order);
        orders.add(order);

        EasyExcel.write("C:/Users/tianlihu/Desktop/test1.xlsx").sheet("测试", Order.class, orders).doWrite();

        List<Order> orders1 = EasyExcel.read("C:/Users/tianlihu/Desktop/test1.xlsx").doReadAll(Order.class);
        orders1.forEach(System.out::println);

        Map<String, Object> map = new HashMap<>();
        map.put("第1列", "南方广利回报债券型证券投资基金A级222");
        map.put("第2列", new Date());
        map.put("第3列", "20230217");
        map.put("第4列", 10);
        map.put("第5列", 3618287.74);
        map.put("第6列", new BigDecimal("3618287.74"));
        map.put("第7列", "3618287.74");
        map.put("第8列", true);

        List<String> headers = Arrays.asList("第1列", "第2列", "第3列", "第4列", "第5列", "第6列", "第7列", "第8列");
        EasyExcel.write("C:/Users/tianlihu/Desktop/test2.xlsx").sheet("测试", headers, Arrays.asList(map, map)).doWrite();
    }
}
