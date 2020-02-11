package com.xingren.excel.converter.write;

import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/2/8 12:28 下午
 */
public class OffSetDateTimeWriteConverter implements IWriteConverter {

    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {
        if (!OffsetDateTime.class.equals(entity.getField().getType())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 OffsetDateTime 类型!");
        }

        Method getMethod = ReflectorUtil.forClass(clazz).getGetMethod(entity.getFiledName());
        OffsetDateTime offsetDateTime = null;
        try {
            offsetDateTime = (OffsetDateTime) getMethod.invoke(rowData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        return DateUtil.formatOffsetDateTime(offsetDateTime, entity.getDatePattern());

    }
}
