package com.xingren.excel.handler.impl;

import com.xingren.excel.ExcelConstant;
import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 默认样式垂直水平居中
 *
 * @author guang
 * @since 2020/2/27 3:34 下午
 */
public class DefaultColumnNameCellStyleHandler implements ICellStyleHandler {

    @Override
    public CellStyle handle(Workbook workbook, CellStyle cellStyle, Object rowData, ExcelColumnAnnoEntity entity) {
        return ExcelConstant.defaultColumnNameStyle(workbook);
    }
}
