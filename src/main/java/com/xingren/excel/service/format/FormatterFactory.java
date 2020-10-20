package com.xingren.excel.service.format;

import com.xingren.excel.enums.FormatType;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.service.format.impl.DropDownListFormatter;

/**
 * @author ellin
 * @since 2020/08/03
 */
public class FormatterFactory {

    public static IFormatter get(FormatType type) {
        switch (type) {
            case DROP_DOWN_LIST:
                return new DropDownListFormatter();
            default:
                throw new ExcelException("无效格式化类型");
        }
    }

}
