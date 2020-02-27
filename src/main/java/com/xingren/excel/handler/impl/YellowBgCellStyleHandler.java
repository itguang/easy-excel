package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 黄色背景
 *
 * @author guang
 * @since 2020/2/27 3:07 下午
 */
public class YellowBgCellStyleHandler implements ICellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, CellStyle style, Object rowData, ExcelColumnAnnoEntity entity) {
        // 填充色
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        return style;
    }
}
