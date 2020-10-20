package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.lang3.StringUtils;

/**
 * @author guang
 * @since 2020/8/17 6:57 下午
 */
public class IntegerReadConverter implements IReadConverter<Object, Integer> {
    @Override
    public Integer convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        Class<?> fieldType = annoEntity.getField().getType();
        String cellValue = cellEntity.getCellValue();

        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }

        if (Integer.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
            if (annoEntity.getYuanToCent()) {
                double doubleValue = Double.parseDouble(cellValue) * 100;
                return (int) doubleValue;
            } else {
                return Integer.valueOf(cellValue);

            }
        }
        return null;
    }
}
