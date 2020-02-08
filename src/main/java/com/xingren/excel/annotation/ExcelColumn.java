package com.xingren.excel.annotation;

import com.xingren.excel.converter.DefaultConverter;
import com.xingren.excel.converter.IConverter;

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
    String columnName();

    /**
     * 排序
     */
    int index();

    /**
     * 日期格式
     */
    String datePattern() default "yyyy/MM/dd";

    /**
     * 枚举导入使用的key
     */
    String enumKey() default "";

    /**
     * true 转换
     *
     */
    String trueToStr() default "true";

    /**
     * false 转换
     *
     */
    String falseToStr() default "false";

    /**
     * 前缀
     */
    String prefix() default "";

    /**
     * 后缀
     */
    String suffix() default "";

    /**
     * 是否启用 分转元 ,仅支持 Integer 和 Long 类型
     */
    boolean centToYuan() default false;

    /**
     * 转换器
     */
    Class<? extends IConverter> converter() default DefaultConverter.class;

}
