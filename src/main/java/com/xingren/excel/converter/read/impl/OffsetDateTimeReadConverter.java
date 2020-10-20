package com.xingren.excel.converter.read.impl;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.DateUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;

import java.time.OffsetDateTime;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * @author guang
 * @since 2020/8/17 7:16 下午
 */
public class OffsetDateTimeReadConverter implements IReadConverter<Object, OffsetDateTime> {
    @Override
    public OffsetDateTime convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, Object rowObj) {
        Class<?> fieldType = annoEntity.getField().getType();
        String cellValue = cellEntity.getCellValue();
        if (StringUtils.isEmpty(cellValue)) {
            return null;
        }
        if (OffsetDateTime.class.equals(fieldType)) {
            if (StringUtils.isNotEmpty(cellValue)) {
                Cell cell = cellEntity.getCell();
                OffsetDateTime offsetDateTime = null;
                if (cell.getCellTypeEnum().equals(NUMERIC)) {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(cellValue, DEFAULT_DATE_PATTREN);
                } else {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(cellValue, annoEntity.getDatePattern());
                }
                return offsetDateTime;
            }
        }
        return null;
    }
}
