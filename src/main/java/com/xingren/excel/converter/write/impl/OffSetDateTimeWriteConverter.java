package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.Method;
import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/2/8 12:28 下午
 */
public class OffSetDateTimeWriteConverter implements IWriteConverter {

    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Object rowData) {
        Class<?> clazz = rowData.getClass();
        if (!OffsetDateTime.class.equals(entity.getField().getType())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 OffsetDateTime 类型!");
        }

        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(entity.getFiledName());

        OffsetDateTime offsetDateTime = (OffsetDateTime) ReflectorUtil.invokeGetMethod(getMethod, rowData);

        return DateUtil.formatOffsetDateTime(offsetDateTime, entity.getDatePattern());

    }
}
