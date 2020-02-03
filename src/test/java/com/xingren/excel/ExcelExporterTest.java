package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import org.apache.commons.lang3.StringUtils;
import org.junit.Before;
import org.junit.Test;

import java.time.OffsetDateTime;
import java.util.ArrayList;

import static org.junit.Assert.assertTrue;

public class ExcelExporterTest {

    ArrayList<Fruit> fruits;

    @Before
    public void before() {
        Fruit apple = new Fruit(1, 1000L, OffsetDateTime.now(), "苹果");
        Fruit orange = new Fruit(2, 2000L, OffsetDateTime.now().minusDays(1), "橘子");

        fruits = new ArrayList<>();
        fruits.add(apple);
        fruits.add(orange);

    }

    @Test
    public void testExport() {
        ExcelExporter.create().export(fruits, Fruit.class);
    }

    @Test
    public void testSheetName() {
        String sheetName = "测试 sheet name";
        ExcelExporter excelExporter = ExcelExporter.create(ExcelType.XLSX)
                .sheetName(sheetName);
        assertTrue(excelExporter.getSheetName().equals(sheetName));
    }

    @Test
    public void testDefaultSheetName() {
        ExcelExporter excelExporter = ExcelExporter.create(ExcelType.XLSX);
        assertTrue(ExcelConstant.DEFAULT_SHEET_NAME.equals(excelExporter.getSheetName()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void testSheetNameIsEmpty() {
        ExcelExporter.create(ExcelType.XLSX).sheetName(null);
    }

    @Test
    public void testSheetheaderIsEmpty() {
        assertTrue(StringUtils.isEmpty(ExcelExporter.create(ExcelType.XLSX).getSheetHeader()));
    }

    @Test
    public void testSheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelExporter excelExporter = ExcelExporter.create(ExcelType.XLSX).sheetheader(sheetHeader);
        assertTrue(excelExporter.getSheetHeader().equals(sheetHeader));
    }

}