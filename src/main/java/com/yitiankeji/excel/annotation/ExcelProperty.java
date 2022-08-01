package com.yitiankeji.excel.annotation;

import com.yitiankeji.excel.converter.Converter;
import com.yitiankeji.excel.converter.Converter.AutoConverter;

import java.lang.annotation.*;

@Target(ElementType.FIELD)
@Retention(RetentionPolicy.RUNTIME)
@Inherited
public @interface ExcelProperty {

    String value();

    int index() default -1;

    String format() default "";

    Class<? extends Converter> converter() default AutoConverter.class;
}
