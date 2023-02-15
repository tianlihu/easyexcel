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

    /** 数据类型(0字符串, 1数字, 2日期) **/
    int type() default 0;

    /** 用于自定义类型转换 **/
    Class<? extends Converter> converter() default AutoConverter.class;

    /** 数据类型 - 字符串 **/
    int STRING = 0;
    /** 数据类型 - 数字 **/
    int NUMBER = 1;
    /** 数据类型 - 日期 **/
    int DATE = 2;
}
