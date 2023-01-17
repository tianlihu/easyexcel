package com.yitiankeji.excel.annotation;

import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.converter.Converter.AutoConverter;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

    /** 列名 **/
    String[] value();

    /** 排序 **/
    int index() default -1;

    /** 格式：支持日期和数字的格式 **/
    String format() default "";

    /** 额外信息，用于支持更多的字典项等 **/
    String extra() default "";

    /** 用于自定义类型转换 **/
    Class<? extends Converter> converter() default AutoConverter.class;
}
