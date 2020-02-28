package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.service.write.ExcelWriteService;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.ArrayList;
import java.util.Collections;
import java.util.List;
import java.util.Map;
import java.util.concurrent.ConcurrentHashMap;

import static com.xingren.excel.ExcelConstant.DEFAULT_SHEET_NAME;
import static com.xingren.excel.ExcelConstant.columnDataRowHeight;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelWriter {

    /**
     * CellStyle 创建数量有限,保证并发安全
     */
    Map<String, CellStyle> cacheStyle = new ConcurrentHashMap<>();

    private Workbook workbook;

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

    private ExcelWriteService excelWriteService;

    private ExcelWriter() {
    }

    private ExcelWriter(ExcelType excelType) {
        this.excelType = excelType;
        if (ExcelType.XLS.equals(excelType)) {
            workbook = new HSSFWorkbook();
        } else {
            workbook = new XSSFWorkbook();
        }
    }

    /**
     * 导出模板
     */
    public <T> Workbook writeTemplate(Class<T> clazz) {
        ArrayList<T> rowDatas = new ArrayList<>();
        return write(rowDatas, clazz);
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
        excelWriteService = ExcelWriteService.forClass(clazz);

        // 创建表格
        Sheet sheet = workbook.createSheet(sheetName);
        sheet.autoSizeColumn(rowIndex);

        // 创建表格头
        if (StringUtils.isNotEmpty(sheetHeader)) {
            excelWriteService.createSheetHeader(workbook, rowIndex, sheetHeader, sheet);
            rowIndex = rowIndex + 1;
        }

        // 创建列标题
        List<ExcelColumnAnnoEntity> annoEntities = ExcelColumnService.forClass(clazz).getOrderedExcelColumnEntity();
        if (CollectionUtils.isNotEmpty(annoEntities)) {
            createColumnTitle(rowIndex, sheet, annoEntities);
        }

        // 创建 rows
        CellStyle style = ExcelConstant.defaultDataRowStyle(workbook);
        for (Object rowData : rows) {
            Row row = sheet.createRow(++rowIndex);
            insertRowData(annoEntities, rowData, row);
            row.setRowStyle(style);
            row.setHeightInPoints(columnDataRowHeight);
        }

        workbook.setActiveSheet(activeSheet);
        return workbook;
    }

    private void insertRowData(List<ExcelColumnAnnoEntity> annoEntities,
                               Object rowData, Row row) {
        for (int i = 0; i < annoEntities.size(); i++) {
            ExcelColumnAnnoEntity entity = annoEntities.get(i);
            Cell cell = row.createCell(i);
            //TODO # guang 待优化
            CellStyle defaultCellStyle = getDefaultCellStyle(entity);
            defaultCellStyle = entity.getCellStyleHandler().handle(workbook, defaultCellStyle, rowData, entity);
            //  Cell 样式设置
            cell.setCellStyle(defaultCellStyle);
            Object value = excelWriteService.parseFieldValue(rowData, entity, cell);
            cell.setCellValue(value == null ? "" : value.toString());

        }
    }

    private CellStyle getDefaultCellStyle(ExcelColumnAnnoEntity entity) {
        CellStyle cellStyle;
        if (cacheStyle.get(entity.getColumnName()) == null) {
            cellStyle = ExcelConstant.defaultDataRowStyle(workbook);
            cacheStyle.put(entity.getColumnName(), cellStyle);
        } else {
            cellStyle = cacheStyle.get(entity.getColumnName());
        }
        return cellStyle;
    }

    /**
     * 创建 ColumnTitle 并设置宽度自适应
     */
    private void createColumnTitle(int rowIndex, Sheet sheet, List<ExcelColumnAnnoEntity> annoEntities) {
        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(ExcelConstant.defaultColumnNameStyle(workbook));
        row.setHeightInPoints(ExcelConstant.columnTitleRowHeight);

        for (int columnNum = 0; columnNum < annoEntities.size(); columnNum++) {
            int columnWidth = sheet.getColumnWidth(columnNum) / 256;
            ExcelColumnAnnoEntity annoEntity = annoEntities.get(columnNum);

            Cell currentCell = row.createCell(columnNum);
            // columnName Style 设置
            CellStyle columnCellStyle = annoEntity.getColumnNameCellStyleHandler().handle(workbook, annoEntity);
            if (null != columnCellStyle) {
                currentCell.setCellStyle(columnCellStyle);
            }
            currentCell.setCellValue(annoEntity.getColumnName());

            // 自动列宽
            int length = currentCell.getStringCellValue().getBytes().length;
            if (columnWidth < length) {
                columnWidth = length;
            }
            sheet.setColumnWidth(columnNum, columnWidth * 256);
        }

    }

    public static ExcelWriter create(ExcelType excelType) {
        return new ExcelWriter(excelType);
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
