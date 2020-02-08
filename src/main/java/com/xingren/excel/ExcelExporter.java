package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.export.ExcelExportService;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;
import lombok.SneakyThrows;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.Collection;
import java.util.Collections;
import java.util.List;

import static com.xingren.excel.ExcelConstant.DEFAULT_SHEET_NAME;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelExporter {

    private static Workbook workbook = new HSSFWorkbook();

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
     * 列头高度
     */
    private int columnTitleHeight = 20;

    /**
     * Excel 文件类型
     */
    private ExcelType excelType = DEFAULT_EXCEL_TYPE;

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
        if (CollectionUtils.isEmpty(rows)) {
            rows = Collections.emptyList();
        }
        if (null == clazz) {
            throw new ExcelException("clazz 不能为空");
        }

        int rowIndex = 0;

        ReflectorUtil reflectorUtil = ReflectorUtil.forClass(clazz);
        ExcelExportService excelExportService = new ExcelExportService();

        // 生成一个表格
        Sheet sheet = workbook.createSheet(sheetName);

        // 设置表格头
        if (StringUtils.isNotEmpty(sheetHeader)) {
            List<Field> excelField = excelExportService.filterExcelField(reflectorUtil.getFieldList());
            if (CollectionUtils.isEmpty(excelField)) {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 0));
            } else {
                sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, excelField.size() - 1));
            }
            Row row0 = sheet.createRow(rowIndex);
            Cell header = row0.createCell(0);
            header.setCellValue(sheetHeader);
            row0.setHeightInPoints(columnTitleHeight * 2);
            rowIndex = rowIndex + 1;
        }

        //设置列标题
        List<ExcelColumnAnnoEntity> annoEntities = excelExportService.getOrderedExcelColumnEntity(clazz);
        if (CollectionUtils.isNotEmpty(annoEntities)) {
            Row row = sheet.createRow(rowIndex);
            for (int i = 0; i < annoEntities.size(); i++) {
                row.createCell(i).setCellValue(annoEntities.get(i).getColumnName());
                // 列标题高度
                row.setHeightInPoints(columnTitleHeight);
            }
        }

        // 设置 rows 数据
        for (Object rowData : rows) {
            Row row = sheet.createRow(++rowIndex);
            for (int i = 0; i < annoEntities.size(); i++) {
                ExcelColumnAnnoEntity entity = annoEntities.get(i);
                Object value = excelExportService.parseFieldValue(clazz, rowData, entity);
                row.createCell(i).setCellValue(value == null ? "" : value.toString());
            }
        }

        workbook.setActiveSheet(activeSheet);
        return workbook;
    }

    @SneakyThrows
    public static ExcelExporter create(ExcelType excelType) {
        if (ExcelType.XLSX.equals(excelType)) {
            workbook = WorkbookFactory.create(true);
        } else {
            workbook = WorkbookFactory.create(false);
        }

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

    public ExcelExporter activeSheet(int activeSheet) {
        if (activeSheet < 0) {
            activeSheet = 0;
        }
        this.activeSheet = activeSheet;
        return this;
    }

    public ExcelExporter sheetheader(String sheetHeader) {
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
