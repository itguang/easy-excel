package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.export.ExcelService;
import com.xingren.excel.pojo.ColumnEntity;
import com.xingren.excel.pojo.RowEntity;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
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
    private int startRowNum = 0;

    /**
     * 读取第几个Sheet
     */
    private int sheetNum = 0;

    private Workbook workbook;

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
     * 从哪一行(包含 columnTitle )开始读取,第一行为 startRowNum = 0
     * <Br/> 默从第0行开始读取,并且第0行为 columnTitle
     */
    public ExcelReader startRowNum(int startRowNum) {
        this.startRowNum = startRowNum;
        return this;
    }

    public static ExcelReader read(InputStream inputStream) {
        return read(inputStream, ExcelType.XLSX);
    }

    public static ExcelReader read(InputStream inputStream, ExcelType excelType) {
        return new ExcelReader(inputStream, excelType);
    }

    public <T> List<T> to(Class<T> clazz) {
        List<T> rowDataList = null;
        Sheet sheet = workbook.getSheetAt(sheetNum);
        int lastRowIndex = sheet.getLastRowNum();
        Row columnTitleRow = sheet.getRow(startRowNum++);

        String[] columnNames = getColumnNames(columnTitleRow);

        ArrayList<RowEntity> rowEntityList = new ArrayList<>();
        for (int rowNum = startRowNum; rowNum <= lastRowIndex; rowNum++) {
            List<ColumnEntity> columnEntityList = new ArrayList<>();
            Row row = sheet.getRow(rowNum);
            for (int curCellNum = row.getFirstCellNum(); curCellNum < row.getLastCellNum(); curCellNum++) {
                String columnName = columnNames[curCellNum];
                String columnValue = row.getCell(curCellNum).getStringCellValue();
                ColumnEntity columnEntity = new ColumnEntity(columnName, columnValue);
                columnEntityList.add(columnEntity);
            }
            rowEntityList.add(new RowEntity(columnEntityList));
        }
        rowDataList = ExcelService.forClass(clazz).parseRowEntity(rowEntityList);

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
