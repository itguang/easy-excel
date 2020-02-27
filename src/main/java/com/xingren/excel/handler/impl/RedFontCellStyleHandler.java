package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 红色字体
 *
 * @author guang
 * @since 2020/2/27 2:57 下午
 */
public class RedFontCellStyleHandler implements ICellStyleHandler {

    @Override
    public CellStyle handle(Workbook workbook, CellStyle style, Object rowData, ExcelColumnAnnoEntity entity) {
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);
        return style;
    }
}
