package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.Method;

/**
 * @author guang
 * @since 2020/8/21 10:46 上午
 */
public class IntegerWriteConverter implements IWriteConverter<Object, Integer> {
    @Override
    public Integer convert(ExcelColumnAnnoEntity entity, Object rowData) {
        Class<?> clazz = rowData.getClass();
        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(entity.getFiledName());
        Object value = ReflectorUtil.invokeGetMethod(getMethod, rowData);
        return (Integer) value;
    }
}
