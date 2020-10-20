package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.Method;
import java.math.BigDecimal;

/**
 * @author guang
 * @since 2020/9/7 3:24 下午
 */
public class BigDecimalWriteConverter implements IWriteConverter<Object, BigDecimal> {
    @Override
    public BigDecimal convert(ExcelColumnAnnoEntity entity, Object rowData) {
        Class<?> clazz = rowData.getClass();
        if (!BigDecimal.class.equals(entity.getField().getType())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 BigDecimal 类型!");
        }
        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(entity.getFiledName());
        BigDecimal bigDecimal = (BigDecimal) ReflectorUtil.invokeGetMethod(getMethod, rowData);
        if (null == bigDecimal) {
            return null;
        }
        return bigDecimal;
    }
}
