package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.lang3.StringUtils;

/**
 * @author guang
 * @since 2020/8/17 6:01 下午
 */
public class BooleanReadConverter implements IReadConverter<Object, Boolean> {
    @Override
    public Boolean convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        String cellValue = cellEntity.getCellValue();

        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }

        String strToFalse = annoEntity.getStrToFalse();
        String strToTrue = annoEntity.getStrToTrue();
        ;
        if (StringUtils.isNotEmpty(strToFalse) && cellValue.equals(strToFalse)
                || StringUtils.equalsIgnoreCase("false", cellValue)) {
            return Boolean.FALSE;
        }
        if (StringUtils.isNotEmpty(strToTrue) && cellValue.equals(strToTrue)
                || StringUtils.equalsIgnoreCase("true", cellValue)) {
            return Boolean.TRUE;
        }
        return null;
    }

}
