package com.xingren.excel.converter.read;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

public interface IReadConverter {

    /**
     * @param entity    Column 注解 实体
     * @param clazz     row class
     * @param cellValue
     * @return 转换后的值
     */
    Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object cellValue);
}
