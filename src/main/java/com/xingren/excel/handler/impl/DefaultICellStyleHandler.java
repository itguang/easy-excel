package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 默认样式垂直水平居中
 *
 * @author guang
 * @since 2020/2/27 2:04 下午
 */
public class DefaultICellStyleHandler implements ICellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, CellStyle style, Object rowData, ExcelColumnAnnoEntity entity) {
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }
}
