package com.xingren.excel.handler;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * 工作簿Workbook中单元格样式个数是有限制的，所以在  程序中应该重复使用相同CellStyle，而不是为每个单元  格创建一个CellStyle
 * CellStyle 处理器
 */
public interface ICellStyleHandler {

    /**
     * <p>
     * workbook 创建样式,并进行设置
     * <code>CellStyle cellStyle = workbook.createCellStyle(); </code>
     * <p>
     * 需要注意的是, xlsx 格式的 excel 文件中 做多只允许创建 64000 个CellStyle
     * </p>
     *
     * @param workbook  工作簿
     * @param cellStyle cellStyle
     * @param rowData   当前行
     * @param entity    当前字段上的  ExcelColumn 注解值
     * @return 返回一个 CellStyle 样式
     */
    CellStyle handle(Workbook workbook, CellStyle cellStyle, Object rowData, ExcelColumnAnnoEntity entity);
}
