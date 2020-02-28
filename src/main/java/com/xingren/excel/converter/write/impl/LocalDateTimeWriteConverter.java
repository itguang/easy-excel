package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;

/**
 * @author guang
 * @since 2020/2/8 3:42 下午
 */
public class LocalDateTimeWriteConverter implements IWriteConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {

        if (!LocalDateTime.class.equals(entity.getField().getType())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 LocalDateTime 类型!");
        }

        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(entity.getFiledName());
        LocalDateTime localDateTime = null;
        try {
            localDateTime = (LocalDateTime) getMethod.invoke(rowData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        return DateUtil.formatLocalDateTime(localDateTime, entity.getDatePattern());
    }
}
