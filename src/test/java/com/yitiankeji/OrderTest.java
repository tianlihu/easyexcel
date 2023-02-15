package com.yitiankeji;

import com.yitiankeji.excel.EasyExcel;
import lombok.SneakyThrows;

import java.math.BigDecimal;
import java.util.ArrayList;
import java.util.Date;
import java.util.List;

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

        EasyExcel.write("C:/Users/tianlihu/Desktop/test.xlsx").sheet("测试", Order.class, orders).doWrite();

        List<Order> orders1 = EasyExcel.read("C:/Users/tianlihu/Desktop/test.xlsx").doReadAll(Order.class);
        orders1.forEach(System.out::println);
    }
}
