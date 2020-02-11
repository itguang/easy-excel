package com.xingren.excel.converter.write;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * Excel 转换器
 */
public interface IWriteConverter {

    /**
     * @param entity  Column 注解 实体
     * @param clazz   row class
     * @param rowData 要转换的对象
     * @return 转换后的值
     */
    Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData);

}
