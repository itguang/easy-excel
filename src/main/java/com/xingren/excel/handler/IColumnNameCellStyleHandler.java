package com.xingren.excel.handler;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.*;

/**
 * ColumnName 样式设置
 */
public interface IColumnNameCellStyleHandler {

    /**
     * 从缓存中创建 style,默认按 columnName 进行缓存,如果需要根据 cell 的值动态创建不同的 style,请自行创建 CellStyle,并注意缓存起来
     *
     * @param workbook
     * @param entity
     * @return
     */
    default CellStyle createCellStyle(Workbook workbook, ExcelColumnAnnoEntity entity) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        Font font = workbook.createFont();
        // 加粗
        font.setBold(true);
        style.setFont(font);
        return style;
    }

    /**
     * <p>
     * workbook 创建样式,并进行设置
     * <code>CellStyle cellStyle = workbook.createCellStyle(); </code>
     * <p>
     * 需要注意的是, xlsx 格式的 excel 文件中 做多只允许创建 64000 个CellStyle
     * </p>
     *
     * @param workbook 工作簿
     * @param entity   当前字段上的  ExcelColumn 注解值
     * @return 返回一个 CellStyle 样式
     */
    CellStyle handle(Workbook workbook, ExcelColumnAnnoEntity entity);
}
