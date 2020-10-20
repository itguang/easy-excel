package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.lang3.StringUtils;

/**
 * @author guang
 * @since 2020/8/17 5:41 下午
 */
public class PrefixAndSuffixReadConverter implements IReadConverter<Object, String> {
    @Override
    public String convert(ExcelColumnAnnoEntity entity, CellEntity cellEntity, Object rowObj) {
        String cellValue = cellEntity.getCellValue();
        if (StringUtils.isEmpty(cellValue)) {
            return cellValue;
        }

        // 读取的时候把前缀后缀去掉
        if (StringUtils.isNotEmpty(entity.getPrefix()) || StringUtils.isNotEmpty(entity.getSuffix())) {
            cellValue = StringUtils.substring(cellValue, entity.getPrefix().length(),
                    entity.getPrefix().length() + cellValue.length() - entity.getSuffix().length());
        }

        return cellValue;
    }
}
