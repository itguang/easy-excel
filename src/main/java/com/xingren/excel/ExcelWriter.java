package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.ErrorInfoRow;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.service.write.ExcelWriteService;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.util.*;
import java.util.concurrent.ConcurrentHashMap;

import static com.xingren.excel.ExcelConstant.*;

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
     * workBook 中所有的 sheetName
     */
    private LinkedList<String> sheetNames = new LinkedList<>();

    /**
     * sheet 个数
     */
    private Integer sheetIndex = 0;

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
        if (ExcelType.XLS.equals(this.excelType)) {
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

    public <T> ExcelWriter addSheetTemplate(String sheetName, String sheetHeader, Class<T> clazz) {
        this.sheetName = sheetName;
        this.sheetHeader = sheetHeader;
        writeTemplate(clazz);
        return this;
    }

    public <T> ExcelWriter addSheetTemplate(String sheetName, Class<T> clazz) {
        this.sheetName = sheetName;
        this.sheetHeader = null;
        writeTemplate(clazz);
        return this;
    }

    public <T> ExcelWriter addSheet(String sheetName, String sheetHeader, List<T> rows, Class<T> clazz) {
        this.sheetName = sheetName;
        this.sheetHeader = sheetHeader;
        write(rows, clazz);
        return this;
    }

    public <T> ExcelWriter addSheet(String sheetName, List<T> rows, Class<T> clazz) {
        this.sheetName = sheetName;
        this.sheetHeader = null;
        write(rows, clazz);
        return this;
    }

    /**
     * 导出excel
     *
     * @param rows  java Entity Collection
     * @param clazz java Entity Class
     * @return workBook
     */
    public <T> Workbook write(List<T> rows, Class<T> clazz) {
        if (CollectionUtils.isEmpty(rows)) {
            rows = Collections.emptyList();
        }
        if (null == clazz) {
            throw new ExcelException("clazz 不能为空");
        }

        int rowIndex = 0;
        excelWriteService = ExcelWriteService.forClass(clazz);

        // 创建表格
        Sheet sheet = createSheet();
        sheet.autoSizeColumn(rowIndex);
        // 是否需要动态删除最后一列(ErrorInfoRow)
        boolean needMinusLastColumn = isNeedMinusLastColumn(rows, clazz);

        // 创建表格头
        if (StringUtils.isNotEmpty(sheetHeader)) {
            excelWriteService.createSheetHeader(workbook, rowIndex, sheetHeader, sheet, needMinusLastColumn);
            rowIndex = rowIndex + 1;
        }

        // 创建列标题
        List<ExcelColumnAnnoEntity> annoEntities = ExcelColumnService.forClass(clazz).getOrderedExcelColumnEntity();
        if (CollectionUtils.isNotEmpty(annoEntities)) {
            createColumnTitle(rowIndex, sheet, annoEntities, needMinusLastColumn);
        }

        // 创建 rows
        CellStyle style = ExcelConstant.defaultDataRowStyle(workbook);
        for (Object rowData : rows) {
            // 增加一行
            Row row = sheet.createRow(++rowIndex);
            insertRowData(annoEntities, rowData, row, needMinusLastColumn, sheet);

            // 样式设置
            row.setRowStyle(style);
            row.setHeightInPoints(columnDataRowHeight);

        }

        workbook.setActiveSheet(activeSheet);
        return workbook;
    }

    /**
     * 创建一个新 sheet
     *
     * @return 新的 sheet 对象
     */
    private Sheet createSheet() {
        sheetIndex = sheetIndex + 1;
        if (StringUtils.isBlank(sheetName)) {
            throw new ExcelException("第 " + sheetIndex + " 个 sheetName 不能为空");
        }
        if (sheetNames.contains(sheetName)) {
            throw new ExcelException("当前 workbook 已经包含了一个 sheet 名为: '" + sheetName + "'");
        }
        sheetNames.add(sheetName);
        return workbook.createSheet(sheetName);
    }

    private <T> boolean isNeedMinusLastColumn(List<T> rows, Class<T> clazz) {

        // 是否继承自 isErrorInfoRow
        boolean isErrorInfoRow = isErrorInfoRow(clazz);

        // 是否有错误信息
        boolean hasErrorInfo = hasErrorInfo(rows, clazz);

        return isErrorInfoRow && !hasErrorInfo;
    }

    private void insertRowData(List<ExcelColumnAnnoEntity> annoEntities,
                               Object rowData, Row row, boolean needMinusLastColumn, Sheet sheet) {
        // 计算有效的列
        int columnSize = calcColumnSize(annoEntities, needMinusLastColumn);
        for (int columnNum = 0; columnNum < columnSize; columnNum++) {
            ExcelColumnAnnoEntity entity = annoEntities.get(columnNum);
            Cell currentCell = row.createCell(columnNum);

            //  Cell 样式设置
            CellStyle defaultCellStyle = getDefaultCellStyle(entity);
            defaultCellStyle = entity.getCellStyleHandler().handle(workbook, defaultCellStyle, rowData, entity);
            currentCell.setCellStyle(defaultCellStyle);
            // 值设置
            Object value = excelWriteService.parseFieldValue(rowData, entity, currentCell);
            String cellValue = value == null ? "" : value.toString();
            currentCell.setCellValue(cellValue);

            // 设置自动列宽
            excelWriteService.autoCellWidth(sheet, columnNum, cellValue);

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
    private <T> void createColumnTitle(int rowIndex, Sheet sheet, List<ExcelColumnAnnoEntity> annoEntities,
                                       boolean needMinusLastColumn) {

        Row row = sheet.createRow(rowIndex);
        row.setRowStyle(ExcelConstant.defaultColumnNameStyle(workbook));
        // ColumnName 高度设置
        row.setHeightInPoints(ExcelConstant.columnTitleRowHeight);

        // 一共有多少有效的列
        int columnSize = calcColumnSize(annoEntities, needMinusLastColumn);
        for (int columnNum = 0; columnNum < columnSize; columnNum++) {
            ExcelColumnAnnoEntity annoEntity = annoEntities.get(columnNum);

            Cell currentCell = row.createCell(columnNum);
            // columnName Style 设置
            CellStyle defaultColumnNameStyle = ExcelConstant.defaultColumnNameStyle(workbook);
            CellStyle columnCellStyle = annoEntity.getColumnNameCellStyleHandler().handle(workbook,
                    defaultColumnNameStyle, null, annoEntity);
            if (null != columnCellStyle) {
                currentCell.setCellStyle(columnCellStyle);
            }

            String columnName = annoEntity.getColumnName();
            // 根据是否是 ErrorInfoRow 类型来选择是否需要 错误信息 列
            if (needMinusLastColumn && ERROR_COLUMN_NAME.equals(columnName)) {
                continue;
            }
            currentCell.setCellValue(columnName);
            // 自动列宽
            excelWriteService.autoCellWidth(sheet, columnNum, columnName);
        }

    }

    /**
     * 计算有效列数
     */
    private int calcColumnSize(List<ExcelColumnAnnoEntity> annoEntities, boolean needMinusLastColumn) {
        int columnSize = annoEntities.size();
        if (needMinusLastColumn) {
            columnSize = columnSize - 1;
        }
        return columnSize;
    }

    private <T> boolean hasErrorInfo(List<T> rows, Class<T> clazz) {
        //是否是 ErrorInfoRow 的子类
        boolean isErrorInfoRow = isErrorInfoRow(clazz);
        if (!isErrorInfoRow) {
            return false;
        }
        List<ErrorInfoRow> errorInfoRows = (List<ErrorInfoRow>) rows;
        return errorInfoRows.stream()
                .map(ErrorInfoRow::getErrorInfo)
                .filter(errorInfo -> StringUtils.isNotEmpty(errorInfo))
                .findFirst().isPresent();
    }

    /**
     * 是否是 ErrorInfoRow 的子类
     */
    private <T> boolean isErrorInfoRow(Class<T> clazz) {
        boolean assignableFrom = ErrorInfoRow.class.isAssignableFrom(clazz);
        return assignableFrom;
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

    public Workbook getWorkbook() {
        return workbook;
    }
}
