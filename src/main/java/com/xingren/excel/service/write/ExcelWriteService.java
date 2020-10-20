package com.xingren.excel.service.write;

import com.xingren.excel.ExcelConstant;
import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.converter.write.WriteConverterFactory;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.POIUtil;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.collections4.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.util.List;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.sheetHeaderRowHeight;

/**
 * @author guang
 * @since 2020/1/19 4:36 下午
 */
public class ExcelWriteService<T> {

    private static final WriteConverterFactory writeConverterFactory = new WriteConverterFactory();

    private final ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ExcelWriteService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.fromCache(clazz);
    }

    public static <T> ExcelWriteService forClass(Class<T> clazz) {
        return new ExcelWriteService(clazz);
    }

    public Object parseFieldValue(Object rowData, ExcelColumnAnnoEntity annoEntity, Cell cell) {

        IWriteConverter converter = writeConverterFactory.getConverter(annoEntity);
        Object value = converter.convert(annoEntity, rowData);

        // 前后缀处理
        if (StringUtils.isNotEmpty(annoEntity.getPrefix())) {
            value = annoEntity.getPrefix() + value;
        }
        // 后缀处理
        if (StringUtils.isNotEmpty(annoEntity.getSuffix())) {
            value = value + annoEntity.getSuffix();
        }
        return value;
    }

    private List<Field> filterExcelField(List<Field> fieldList) {
        // 过滤出 注解 @ExcelColumn 的字段
        return fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class)
                ).collect(Collectors.toList());
    }

    public void createSheetHeader(Workbook workbook, int rowIndex, String sheetHeader, Sheet sheet,
                                  boolean needMinusLastColumn) {
        List<Field> excelFields = filterExcelField(reflectorUtil.getFieldList());

        if (CollectionUtils.isEmpty(excelFields)) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 0));
        } else {
            int lastCol = excelFields.size() - 1;
            if (needMinusLastColumn) {
                lastCol = lastCol - 1;
            }
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, lastCol));
        }
        Row row0 = sheet.createRow(rowIndex);
        Cell header = row0.createCell(0);
        header.setCellValue(sheetHeader);
        header.setCellStyle(ExcelConstant.defaultHeaderRowStyle(workbook));
        row0.setHeightInPoints(sheetHeaderRowHeight);
        // 自动列宽
        autoCellWidth(sheet, rowIndex, sheetHeader);
    }

    /**
     * 设置自动列宽
     */
    public void autoCellWidth(Sheet sheet, int columnNum, String cellValue) {
        // 自动列宽
        int columnWidth = sheet.getColumnWidth(columnNum) / 256;
        int length = cellValue.getBytes().length;
        if (columnWidth < length) {
            columnWidth = length;
        }
        sheet.setColumnWidth(columnNum, columnWidth * 256);
    }

    public void addSheetTail(Workbook workbook, int lastRowIndex, String sheetTail, Sheet sheet,
                             boolean needMinusLastColumn, Float tailHeight) {
        int insertRowIndex = lastRowIndex + 1;
        List<Field> excelFields = filterExcelField(reflectorUtil.getFieldList());
        if (CollectionUtils.isEmpty(excelFields)) {
            sheet.addMergedRegion(new CellRangeAddress(insertRowIndex, insertRowIndex, 0, 0));
        } else {
            int lastCol = excelFields.size() - 1;
            if (needMinusLastColumn) {
                lastCol = lastCol - 1;
            }
            sheet.addMergedRegion(new CellRangeAddress(insertRowIndex, insertRowIndex, 0, lastCol));
        }

        Row row0 = sheet.createRow(insertRowIndex);
        Cell header = row0.createCell(0);
        header.setCellValue(sheetTail);
        header.setCellStyle(ExcelConstant.defaultTailRowStyle(workbook));

        Float hieght = POIUtil.getExcelCellAutoHeight(sheetTail + "", tailHeight);
        row0.setHeightInPoints(hieght);
    }
}