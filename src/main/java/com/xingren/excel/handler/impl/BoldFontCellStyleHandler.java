package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 加粗
 *
 * @author guang
 * @since 2020/2/27 3:20 下午
 */
public class BoldFontCellStyleHandler implements ICellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, CellStyle style, Object rowData, ExcelColumnAnnoEntity entity) {
        Font font = workbook.createFont();
        // 加粗
        font.setBold(true);
        style.setFont(font);
        return style;
    }
}
