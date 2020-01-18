package com.xingren.annotation;

import com.xingren.converter.DefaultConverter;
import com.xingren.converter.IConverter;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

/**
 * @author guang
 */
@Retention(RetentionPolicy.RUNTIME)
@Target(ElementType.FIELD)
public @interface ExcelColumn {

    /**
     * Excel 每一列的名称
     */
    String name();

    /**
     * 排序
     */
    int order();

    /**
     * 日期格式
     */
    String datePattern();

    /**
     * 枚举导入使用的函数
     */
    String enumImportMethod() default "";

    /**
     * 转换器
     */
    Class<? extends IConverter> converter() default DefaultConverter.class;

}
