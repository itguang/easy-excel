package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.lang3.StringUtils;

/**
 * @author guang
 * @since 2020/8/17 7:01 下午
 */
public class LongReadConverter implements IReadConverter<Object, Long> {
    @Override
    public Long convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        Class<?> fieldType = annoEntity.getField().getType();
        String cellValue = cellEntity.getCellValue();
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        if (Long.class.equals(fieldType)) {
            if (annoEntity.getYuanToCent()) {
                double doubleValue = Double.parseDouble(cellValue) * 100;
                return (long) doubleValue;
            } else {
                return Long.valueOf(cellValue);
            }
        }
        return null;
    }
}
