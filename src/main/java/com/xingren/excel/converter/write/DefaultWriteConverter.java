package com.xingren.excel.converter.write;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * @author guang
 * @since 2020/1/18 9:50 上午
 */
public class DefaultWriteConverter implements IWriteConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {
        throw new UnsupportedOperationException("该类属于注解默认参数,不可实例化");
    }
}
