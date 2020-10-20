package com.xingren.excel.converter.write;

import com.xingren.excel.converter.ConverterFactory;
import com.xingren.excel.converter.write.impl.*;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

import java.math.BigDecimal;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;

/**
 * 获取 IWriteConverter 的工厂
 *
 * @author guang
 * @since 2020/8/20 3:03 下午
 */
public class WriteConverterFactory implements ConverterFactory {
    private static StringWriteConverter stringWriteConverter = new StringWriteConverter();
    private static Cent2YuanWriteConverter cent2YuanWriteConverter = new Cent2YuanWriteConverter();
    private static IntegerWriteConverter integerWriteConverter = new IntegerWriteConverter();
    private static LongWriteConverter longWriteConverter = new LongWriteConverter();
    private static EnumWriteConverter enumWriteConverter = new EnumWriteConverter();
    private static OffSetDateTimeWriteConverter offSetDateTimeWriteConverter = new OffSetDateTimeWriteConverter();
    private static LocalDateTimeWriteConverter localDateTimeWriteConverter = new LocalDateTimeWriteConverter();
    private static BooleanWriteConverter booleanWriteConverter = new BooleanWriteConverter();
    private static BigDecimalWriteConverter bigDecimalConverter = new BigDecimalWriteConverter();

    @Override
    public IWriteConverter getConverter(ExcelColumnAnnoEntity annoEntity) {
        Class<?> type = annoEntity.getField().getType();

        // 自定义的转换器(放在最前面)
        if (null != annoEntity.getWriteConverter()) {
            return annoEntity.getWriteConverter();
        }

        // String
        if (String.class.equals(type)) {
            return stringWriteConverter;
        }

        // Integer 和 Long 类型的 分转元处理
        if (annoEntity.getCentToYuan()) {
            return cent2YuanWriteConverter;
        }

        // Integer
        if (Integer.class.equals(type)) {
            return integerWriteConverter;
        }

        // Long
        if (Long.class.equals(type)) {
            return longWriteConverter;
        }

        // Enum 类型处理
        if (type.isEnum()) {
            return enumWriteConverter;
        }
        // OffsetDateTime 类型处理
        if (OffsetDateTime.class.equals(type)) {
            return offSetDateTimeWriteConverter;
        }
        // LocalDateTime 类型处理
        if (LocalDateTime.class.equals(type)) {
            return localDateTimeWriteConverter;
        }
        // Boolean 类型处理
        if (Boolean.class.equals(type)) {
            return booleanWriteConverter;
        }
        if (BigDecimal.class.equals(type)){
            return bigDecimalConverter;
        }

        throw new ExcelException("不支持的类型转换: " + type);
    }
}
