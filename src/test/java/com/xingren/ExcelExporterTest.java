package com.xingren;

import org.apache.commons.lang3.StringUtils;
import org.junit.Before;
import org.junit.Test;

import static org.junit.Assert.assertTrue;

public class ExcelExporterTest {

    @Before
    public void before() {

    }

    @Test
    public void testSheetName() {
        String sheetName = "测试 sheet name";
        ExcelExporter excelExporter = ExcelExporter.create()
                .sheetName(sheetName);
        assertTrue(excelExporter.getSheetName().equals(sheetName));
    }

    @Test
    public void testDefaultSheetName() {
        ExcelExporter excelExporter = ExcelExporter.create();
        assertTrue(ExcelConstant.DEFAULT_SHEET_NAME.equals(excelExporter.getSheetName()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void testSheetNameIsEmpty() {
        ExcelExporter.create().sheetName(null);
    }

    @Test
    public void testSheetheaderIsEmpty() {
        assertTrue(StringUtils.isEmpty(ExcelExporter.create().getSheetHeader()));
    }

    @Test
    public void testSheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelExporter excelExporter = ExcelExporter.create().sheetheader(sheetHeader);
        assertTrue(excelExporter.getSheetHeader().equals(sheetHeader));
    }



}