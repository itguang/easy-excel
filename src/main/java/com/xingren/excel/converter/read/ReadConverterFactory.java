package com.xingren.excel.converter.read;

import com.xingren.excel.converter.ConverterFactory;
import com.xingren.excel.converter.read.impl.*;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;

/**
 * 获取 IReadConverter 的工厂
 *
 * @author guang
 * @since 2020/8/20 2:27 下午
 */
public class ReadConverterFactory implements ConverterFactory {

    /**
     * 预先创建好所有的 Converter 提高性能
     */
    private static StringReadConverter stringReadConverter = new StringReadConverter();
    private static BooleanReadConverter booleanReadConverter = new BooleanReadConverter();
    private static IntegerReadConverter integerReadConverter = new IntegerReadConverter();
    private static LongReadConverter longReadConverter = new LongReadConverter();
    private static OffsetDateTimeReadConverter offsetDateTimeReadConverter = new OffsetDateTimeReadConverter();
    private static LocalDateTimeReadConverter localDateTimeReadConverter = new LocalDateTimeReadConverter();
    private static EnumReadConverter enumReadConverter = new EnumReadConverter();

    @Override
    public IReadConverter getConverter(ExcelColumnAnnoEntity annoEntity) {
        Class<?> fieldType = annoEntity.getField().getType();

        // 自定义处理器放在最前面
        if (null != annoEntity.getReadConverter()) {
            return annoEntity.getReadConverter();
        }

        // String 类型处理,(不做任何处理,直接反射赋值)
        if (String.class.equals(fieldType)) {
            return stringReadConverter;
        }

        // Boolean
        if (Boolean.class.equals(fieldType)) {
            return booleanReadConverter;
        }

        // Integer 类型处理
        if (Integer.class.equals(fieldType)) {
            return integerReadConverter;
        }

        //  Long 类型处理
        if (Long.class.equals(fieldType)) {
            return longReadConverter;
        }

        // OffsetDateTime 处理
        if (OffsetDateTime.class.equals(fieldType)) {
            return offsetDateTimeReadConverter;
        }
        // LocalDateTime 处理
        if (LocalDateTime.class.equals(fieldType)) {
            return localDateTimeReadConverter;
        }

        // 枚举处理
        if (fieldType.isEnum()) {
            return enumReadConverter;
        }

        throw new ExcelException("不支持的类型转换: " + fieldType + " ,请自定义转换器");
    }

}
