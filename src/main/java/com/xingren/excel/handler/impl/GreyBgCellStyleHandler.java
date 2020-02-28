package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 灰色背景
 *
 * @author guang
 * @since 2020/2/28 3:36 下午
 */
public class GreyBgCellStyleHandler implements ICellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, CellStyle cellStyle, Object rowData, ExcelColumnAnnoEntity entity) {
        // 填充色
        cellStyle.setFillForegroundColor(IndexedColors.GREY_25_PERCENT.getIndex());
        cellStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return cellStyle;
    }
}
