package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @author guang
 * @since 2020/2/8 12:46 下午
 */
public class Cent2YuanWriteConverter implements IWriteConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {

        Class<?> type = entity.getField().getType();

        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(entity.getFiledName());
        Object value = null;
        try {
            value = getMethod.invoke(rowData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        if (null == value) {
            return 0;
        }

        if (Integer.class.equals(type)) {
            double yuan = ((Integer) value).doubleValue() / 100;
            return String.format("%.2f", yuan);

        }
        if (Long.class.equals(type)) {
            double yuan = ((Long) value).doubleValue() / 100;
            return String.format("%.2f", yuan);
        }

        throw new ExcelConvertException("只支持 Integer 和 Long 类型的 分转元操作");

    }
}
