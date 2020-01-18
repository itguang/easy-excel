package com.xingren;

import org.apache.commons.lang3.StringUtils;

import static com.xingren.ExcelConstant.DEFAULT_SHEET_NAME;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelExporter {

    /**
     * sheetName
     */
    private String sheetName = DEFAULT_SHEET_NAME;

    /**
     * 表格头
     */
    private String sheetHeader;

    public static ExcelExporter create() {
        return new ExcelExporter();
    }

    public ExcelExporter sheetName(String sheetName) {
        if (StringUtils.isEmpty(sheetName)) {
            throw new IllegalArgumentException("sheetName cannot be empty");
        }
        this.sheetName = sheetName;
        return this;
    }

    public ExcelExporter sheetheader(String sheetHeader) {
        if (StringUtils.isNotEmpty(sheetHeader)) {
            this.sheetHeader = sheetHeader;
        }
        return this;
    }

    public String getSheetName() {
        return sheetName;
    }

    public String getSheetHeader() {
        return sheetHeader;
    }
}
