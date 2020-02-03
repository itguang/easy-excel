package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;

import java.util.Collection;

import static com.xingren.excel.ExcelConstant.DEFAULT_SHEET_NAME;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelExporter {

    private static final ExcelType DEAFULT_EXCEL_TYPE = ExcelType.XLSX;

    /**
     * sheetName
     */
    private String sheetName = DEFAULT_SHEET_NAME;

    /**
     * 表格头
     */
    private String sheetHeader;

    /**
     * Excel 文件类型
     */
    private ExcelType excelType = DEAFULT_EXCEL_TYPE;

    private ExcelExporter() {
    }

    /**
     * 导出excel
     *
     * @param rows  java Entity Collection
     * @param clazz java Entity Class
     * @return workBook
     */
    public Workbook export(Collection<?> rows, Class<?> clazz) {
        Workbook workbook = null;

        //todo#guang

        return workbook;
    }

    public static ExcelExporter create(ExcelType excelType) {
        return new ExcelExporter();
    }

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
