package com.xingren.excel;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * @author guang
 * @since 2020/2/11 4:24 下午
 */
public class ToStringConverter implements IWriteConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {
        Product product = (Product) rowData;
        return product.toString();
    }
}
