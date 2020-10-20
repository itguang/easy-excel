package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * @author guang
 * @since 2020/8/20 2:33 下午
 */
public class StringReadConverter implements IReadConverter<Object, String> {
    @Override
    public String convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        String cellValue = cellEntity.getCellValue();
        return cellValue;
    }
}
