package com.xingren.excel.handler;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Workbook;

/**
 * CellStyle 处理器
 *
 * <note>工作簿 Workbook 中单元格样式个数是有限制的，所以在  程序中应该重复使用相同 CellStyle，而不是为每个单元  格创建一个CellStyle</note>
 *
 * @author guang
 */
public interface ICellStyleHandler {

    /**
     * 处理样式
     * <note>需要注意的是, xlsx 格式的 excel 文件中 做多只允许创建 64000 个 CellStyle</note>
     *
     * @param workbook  工作簿对象
     * @param cellStyle cellStyle 默认样式垂直水平居中,可在此基础上进行样式调整
     * @param rowData   当前行对象
     * @param entity    当前字段上的  ExcelColumn 注解值
     * @return 返回一个 CellStyle 样式,并应用到当前 Cell 上
     */
    CellStyle handle(Workbook workbook, CellStyle cellStyle, Object rowData, ExcelColumnAnnoEntity entity);
}
