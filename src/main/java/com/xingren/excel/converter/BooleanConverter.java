package com.xingren.excel.converter;

import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;

/**
 * @author guang
 * @since 2020/2/8 4:51 下午
 */
public class BooleanConverter implements IConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {

        if (!Boolean.class.equals(entity.getField().getType())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 Boolean 类型!");
        }

        Method getMethod = ReflectorUtil.forClass(clazz).getGetMethod(entity.getFiledName());
        Boolean b = null;
        try {
            b = (Boolean) getMethod.invoke(rowData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        if (null == b) {
            return null;
        }

        if (b) {
            return entity.getTrueStr();
        } else {
            return entity.getFalseStr();
        }

    }
}
