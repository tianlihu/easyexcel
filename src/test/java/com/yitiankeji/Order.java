package com.yitiankeji;

import com.yitiankeji.excel.annotation.ExcelProperty;
import lombok.Data;

import java.math.BigDecimal;
import java.util.Date;

import static com.yitiankeji.excel.annotation.ExcelProperty.DATE;
import static com.yitiankeji.excel.annotation.ExcelProperty.NUMBER;

@Data
public class Order {

    @ExcelProperty("field1")
    private String field1;
    @ExcelProperty(value = "field2", type = DATE, format = "yyyyMMdd")
    private Date field2;
    @ExcelProperty(value = "field3", type = DATE, format = "yyyyMMdd")
    private String field3;
    @ExcelProperty(value = "field4", type = NUMBER, format = "###,##0")
    private Integer field4;
    @ExcelProperty(value = "field5", type = NUMBER, format = "###,###.00")
    private Double field5;
    @ExcelProperty(value = "field6", type = NUMBER, format = "###,###.00")
    private BigDecimal field6;
    @ExcelProperty(value = "field7", type = NUMBER, format = "###,###.00")
    private String field7;
    @ExcelProperty(value = "field8")
    private boolean field8;
}
