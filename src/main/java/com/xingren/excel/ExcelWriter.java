package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.export.ExcelService;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.Collections;
import java.util.List;

import static com.xingren.excel.ExcelConstant.DEFAULT_SHEET_NAME;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelWriter {

    private static Workbook workbook;

    private static final ExcelType DEFAULT_EXCEL_TYPE = ExcelType.XLSX;

    /**
     * sheetName
     */
    private String sheetName = DEFAULT_SHEET_NAME;

    /**
     * 默认激活第一个 sheet
     */
    private int activeSheet = 0;

    /**
     * 表格头
     */
    private String sheetHeader;

    /**
     * Excel 文件类型
     */
    private ExcelType excelType = DEFAULT_EXCEL_TYPE;

    private ExcelService excelService;

    private ExcelWriter() {
    }

    /**
     * 导出excel
     *
     * @param rows  java Entity Collection
     * @param clazz java Entity Class
     * @return workBook
     */
    public <T> Workbook write(List<T> rows, Class<?> clazz) {
        if (CollectionUtils.isEmpty(rows)) {
            rows = Collections.emptyList();
        }
        if (null == clazz) {
            throw new ExcelException("clazz 不能为空");
        }

        int rowIndex = 0;
        excelService = ExcelService.forClass(clazz);

        // 创建表格
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.autoSizeColumn(rowIndex);

        // 创建表格头
        if (StringUtils.isNotEmpty(sheetHeader)) {
            excelService.createSheetHeader(workbook, rowIndex, sheetHeader, sheet);
            rowIndex = rowIndex + 1;
        }

        // 创建列标题
        List<ExcelColumnAnnoEntity> annoEntities = excelService.getOrderedExcelColumnEntity();
        if (CollectionUtils.isNotEmpty(annoEntities)) {
            createColumnTitle(rowIndex, sheet, annoEntities);
        }

        // 创建 rows
        for (Object rowData : rows) {
            Row row = sheet.createRow(++rowIndex);
            insertRowData(annoEntities, rowData, row);
            row.setRowStyle(ExcelConstant.defaultColumnStyle(workbook));
        }

        workbook.setActiveSheet(activeSheet);
        return workbook;
    }

    private void insertRowData(List<ExcelColumnAnnoEntity> annoEntities,
                               Object rowData, Row row) {
        for (int i = 0; i < annoEntities.size(); i++) {
            ExcelColumnAnnoEntity entity = annoEntities.get(i);
            Object value = excelService.parseFieldValue(rowData, entity);
            row.createCell(i).setCellValue(value == null ? "" : value.toString());
        }
    }

    /**
     * 创建 ColumnTitle 并设置宽度自适应
     */
    private void createColumnTitle(int rowIndex, Sheet sheet, List<ExcelColumnAnnoEntity> annoEntities) {
        Row row = sheet.createRow(rowIndex);
        for (int columnNum = 0; columnNum < annoEntities.size(); columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            Cell currentCell = row.createCell(columnNum);
            currentCell.setCellValue(annoEntities.get(columnNum).getColumnName());
            int length = currentCell.getStringCellValue().getBytes().length;
            if (columnWidth < length) {
                columnWidth = length;
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }
        row.setRowStyle(ExcelConstant.defaultRowTitleStyle(workbook));
        row.setHeightInPoints(ExcelConstant.columnTitleHeight);
    }

    public static ExcelWriter create(ExcelType excelType) {
        if (ExcelType.XLS.equals(excelType)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }

        return new ExcelWriter();
    }

    public static ExcelWriter create() {
        return create(ExcelType.XLSX);
    }

    public ExcelWriter sheetName(String sheetName) {
        if (StringUtils.isEmpty(sheetName)) {
            throw new IllegalArgumentException("sheetName cannot be empty");
        }
        this.sheetName = sheetName;
        return this;
    }

    public ExcelWriter activeSheet(int activeSheet) {
        if (activeSheet < 0) {
            activeSheet = 0;
        }
        this.activeSheet = activeSheet;
        return this;
    }

    public ExcelWriter sheetHeader(String sheetHeader) {
        if (StringUtils.isEmpty(sheetHeader)) {
            throw new ExcelException("sheetHeader 不能为空");
        }

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
