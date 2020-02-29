package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.ColumnEntity;
import com.xingren.excel.pojo.RowEntity;
import com.xingren.excel.service.read.ExcelReadService;
import com.xingren.excel.service.read.MyDataFormatter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public class ExcelReader {

    /**
     * 从第几行开始读取(一般用来跳过 sheetHeader )
     */
    private int columnNameRowNum = 0;

    /**
     * 读取第几个Sheet
     */
    private int sheetNum = 0;

    private Workbook workbook;

    private DataFormatter formatter = new MyDataFormatter();

    private ExcelReader(InputStream inputStream, ExcelType excelType) {
        try {
            if (ExcelType.XLS.equals(excelType)) {
                workbook = new HSSFWorkbook(inputStream);
            } else {
                workbook = new XSSFWorkbook(inputStream);
            }
        } catch (IOException e) {
            e.printStackTrace();
        }
    }

    /**
     * columnName 所在行
     */
    public ExcelReader columnNameRowNum(int columnNameRowNum) {
        this.columnNameRowNum = columnNameRowNum;
        return this;
    }

    public static ExcelReader read(InputStream inputStream) {
        return read(inputStream, ExcelType.XLSX);
    }

    public static ExcelReader read(InputStream inputStream, ExcelType excelType) {
        return new ExcelReader(inputStream, excelType);
    }

    public <T> List<T> toPojo(Class<T> clazz) {
        List<T> rowDataList = null;
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int lastRowIndex = sheet.getLastRowNum();
        Row columnTitleRow = sheet.getRow(columnNameRowNum++);
        Row lastRow = sheet.getRow(lastRowIndex);
        if (columnTitleRow.getLastCellNum() < lastRow.getLastCellNum()) {
            throw new ExcelException("解析 Excel 失败,检查 columnNameRowNum 是否设置正确");
        }

        String[] columnNames = getColumnNames(columnTitleRow);

        ArrayList<RowEntity> rowEntityList = new ArrayList<>();
        for (int rowNum = columnNameRowNum; rowNum <= lastRowIndex; rowNum++) {
            List<ColumnEntity> columnEntityList = new ArrayList<>();
            Row row = sheet.getRow(rowNum);
            for (int curCellNum = row.getFirstCellNum(); curCellNum < row.getLastCellNum(); curCellNum++) {
                String columnName = columnNames[curCellNum];
                Cell cell = row.getCell(curCellNum);
                String columnValue = formatter.formatCellValue(cell);
                ColumnEntity columnEntity = new ColumnEntity(cell, columnName, columnValue);
                columnEntityList.add(columnEntity);
            }
            rowEntityList.add(new RowEntity(columnEntityList));
        }
        rowDataList = ExcelReadService.forClass(clazz).parseRowEntity(rowEntityList);

        return rowDataList;
    }

    private String[] getColumnNames(Row columnTitleRow) {
        String[] columnNames = new String[columnTitleRow.getLastCellNum() + 1];
        for (int cellNum = columnTitleRow.getFirstCellNum(); cellNum < columnTitleRow.getLastCellNum(); cellNum++) {
            String columnName = columnTitleRow.getCell(cellNum).getStringCellValue();
            columnNames[cellNum] = columnName;
        }
        return columnNames;
    }

}
