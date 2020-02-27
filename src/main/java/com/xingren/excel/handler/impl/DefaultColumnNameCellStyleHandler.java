package com.xingren.excel.handler.impl;

import com.xingren.excel.handler.IColumnNameCellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 默认样式垂直水平居中
 *
 * @author guang
 * @since 2020/2/27 3:34 下午
 */
public class DefaultColumnNameCellStyleHandler implements IColumnNameCellStyleHandler {
    @Override
    public CellStyle handle(Workbook workbook, ExcelColumnAnnoEntity entity) {
        return createCellStyle(workbook, entity);
    }
}
