package com.xingren.excel.pojo;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.handler.impl.RedFontCellStyleHandler;

import static com.xingren.excel.ExcelConstant.ERROR_COLUMN_NAME;

/**
 * 错误信息接口,如果继承此类,在导出时,处理每一行的出错信息,发生异常,都不会中断程序,而是会把错误信息,放在 errorInfo 字段中
 *
 * @author guang
 */
public class ErrorInfoRow {

    @ExcelColumn(index = Integer.MAX_VALUE, columnName = ERROR_COLUMN_NAME, columnNameCellStyleHandler =
            RedFontCellStyleHandler.class, cellStyleHandler = RedFontCellStyleHandler.class)
    String errorInfo = "";

    /**
     * 获取 错误信息对象
     *
     * @return String 对象
     */
    public String getErrorInfo() {
        return errorInfo;
    }

    public void setErrorInfo(String errorInfo) {
        this.errorInfo = errorInfo;
    }
}
