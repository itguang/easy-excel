package com.xingren.excel.annotation;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.converter.read.impl.DefaultReadConverter;
import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.converter.write.impl.DefaultWriteConverter;
import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.handler.impl.DefaultColumnNameCellStyleHandler;
import com.xingren.excel.handler.impl.DefaultICellStyleHandler;

import java.lang.annotation.ElementType;
import java.lang.annotation.Retention;
import java.lang.annotation.RetentionPolicy;
import java.lang.annotation.Target;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;

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
     * 排序(仅对导出有效)
     */
    int index();

    /**
     * 日期格式
     */
    String datePattern() default DEFAULT_DATE_PATTREN;

    /**
     * 枚举导入使用的key
     */
    String enumKey() default "";

    /**
     * true 转换(仅对导出有效)
     */
    String trueToStr() default "true";

    /**
     * 字符串转 true(仅对导入有效)
     */
    String strToTrue() default "";

    /**
     * false 转换(仅对导出有效)
     */
    String falseToStr() default "false";

    /**
     * 字符串 转 false(仅对导入有效)
     */
    String strToFalse() default "";

    /**
     * 前缀
     */
    String prefix() default "";

    /**
     * 后缀
     */
    String suffix() default "";

    /**
     * 是否启用 分转元(仅对导出有效) ,仅支持 Integer 和 Long 类型
     */
    boolean centToYuan() default false;

    /**
     * 是否启用 元转分(仅对导入有效),仅支持 Integer 和 Long 类型
     */
    boolean yuanToCent() default false;

    /**
     * 是否必须 (导入使用)
     */
    boolean required() default false;

    /**
     * 导出转换器
     */
    Class<? extends IWriteConverter> writeConverter() default DefaultWriteConverter.class;

    /**
     * 导入转换器
     */
    Class<? extends IReadConverter> readConverter() default DefaultReadConverter.class;

    /**
     * 单个 Data Cell 样式处理器
     */
    Class<? extends ICellStyleHandler> cellStyleHandler() default DefaultICellStyleHandler.class;

    /**
     * columnName 样式处理器
     */
    Class<? extends ICellStyleHandler> columnNameCellStyleHandler() default DefaultColumnNameCellStyleHandler.class;

}
