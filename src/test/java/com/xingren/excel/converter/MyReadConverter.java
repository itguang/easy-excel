package com.xingren.excel.converter;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * @author guang
 * @since 2020/2/11 4:30 下午
 */
public class MyReadConverter implements IReadConverter {
    @Override
    public Object convert(ExcelColumnAnnoEntity entity, CellEntity cellEntity, Object rowObj) {
        String cellValue = cellEntity.getCellValue();
        return "[" + cellValue + "]";
    }
}
