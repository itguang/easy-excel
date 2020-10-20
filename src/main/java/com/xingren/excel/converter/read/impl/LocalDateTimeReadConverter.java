package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.DateUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.time.LocalDateTime;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * @author guang
 * @since 2020/8/17 7:46 下午
 */
public class LocalDateTimeReadConverter implements IReadConverter<Object, LocalDateTime> {
    @Override
    public LocalDateTime convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        Class<?> fieldType = annoEntity.getField().getType();
        String cellValue = cellEntity.getCellValue();
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }

        // LocalDateTime 处理
        if (LocalDateTime.class.equals(fieldType)) {
            Cell cell = cellEntity.getCell();
            LocalDateTime localDateTime = null;
            if (cell.getCellTypeEnum().equals(NUMERIC)) {
                localDateTime = DateUtil.parseToLocalDateTime(cellValue, DEFAULT_DATE_PATTREN);
            } else {
                localDateTime = DateUtil.parseToLocalDateTime(cellValue, annoEntity.getDatePattern());
            }
            return localDateTime;
        }
        return null;
    }
}
