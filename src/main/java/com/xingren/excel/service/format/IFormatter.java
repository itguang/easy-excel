package com.xingren.excel.service.format;

import com.xingren.excel.pojo.TemplateRowFormat;
import org.apache.poi.ss.usermodel.Sheet;

public interface IFormatter {

    /**
     * 单元格格式化
     *
     * @param sheet  当前sheet
     * @param format 当前格式化数据
     */
    void format(Sheet sheet, TemplateRowFormat format);

}
