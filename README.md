# EasyExcel

## Project Build
```
mvn clean package
```

### Entity Class(Demo Class)
```
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
```

### Read Excel Files(*.xls, *.xlsx)
```
// Read the entire file
List<Order> orders1 = EasyExcel.read("filepath.xlsx").doReadAll(Order.class);

// Read only one sheet file
List<Order> orders2 = EasyExcel.read("filepath.xlsx").sheet(0, Order.class).doRead();
```

### Write Excels(*.xlsx)
```
List<Order> orders = new ArrayList<>();

// Write only one sheet file
EasyExcel.write("filepath.xlsx").sheet("Test Sheet", Order.class, orders).doWrite();

// Write multiple sheet file
EasyExcel.write("filepath.xlsx").sheet("Test Sheet1", Order.class, orders).sheet("Test Shee2", Order.class, orders).doWrite();

// White excel file with row data of type: List<Map<String, Object>>
List<Map<String, Object>> datas = new ArrayList<>();
List<String> headers = Arrays.asList("Column 1", "Column 2", "Column 3", "Column 4");
EasyExcel.write("filepath.xlsx").sheet("测试", headers, datas).doWrite();
```
