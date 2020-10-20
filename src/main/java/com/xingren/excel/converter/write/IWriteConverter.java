package com.xingren.excel.converter.write;

import com.xingren.excel.converter.Converter;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * Excel 转换器
 *
 * @param <T> rowObject 对应的 java 类型
 * @param <D> 要转成的目标类型
 */
public interface IWriteConverter<T, D> extends Converter {

    /**
     * @param entity  Column 注解 实体
     * @param rowData 要转换的对象
     * @return 转换后的值
     */
    D convert(ExcelColumnAnnoEntity entity, T rowData);

}
