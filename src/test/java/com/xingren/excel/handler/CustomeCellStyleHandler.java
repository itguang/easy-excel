package com.xingren.excel.handler;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.*;

/**
 * @author guang
 * @since 2020/2/27 2:36 下午
 */
public class CustomeCellStyleHandler implements ICellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, CellStyle style, Object rowData, ExcelColumnAnnoEntity entity) {
        // 填充色
        style.setFillForegroundColor(IndexedColors.YELLOW.getIndex());
        style.setFillPattern(FillPatternType.SOLID_FOREGROUND);

        // 字体颜色
        Font font = workbook.createFont();
        font.setColor(IndexedColors.RED.getIndex());
        style.setFont(font);

        return style;
    }
}
