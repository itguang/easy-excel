package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.export.ExcelExportService;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.Before;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.*;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertTrue;

public class ExcelExporterTest {

    ArrayList<Product> products;
    String testResourcesPath;
    String desktopFile;

    @Before
    public void before() throws FileNotFoundException {

        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String desktop = fsv.getHomeDirectory().getPath();
        desktopFile = desktop + "/product.xls";

        String testResourcesPath = ExcelExporterTest.class.getResource("/").getPath();
        Product apple = new Product(1000,
                1000L, OffsetDateTime.now(),
                "苹果", true, StateEnum.DOWN, LocalDateTime.now(),
                "好吃");
        Product orange = new Product(2000,
                2000L, OffsetDateTime.now().minusDays(1),
                "橘子", false, StateEnum.UP,
                LocalDateTime.now().minusDays(1), "好酸");

        products = new ArrayList<>();
        products.add(apple);
        products.add(orange);

    }

    @Test
    public void testExport() throws IOException {

        Workbook workbook = ExcelExporter.create()
                .sheetName("商品数据")
                .sheetheader("--2月份商品数据--")
                .export(products, Product.class);

        File file = new File(desktopFile);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    public void testGetExcelRowNames() {
        ExcelExportService excelExportService = new ExcelExportService();
        List<ExcelColumnAnnoEntity> excelRowNames = excelExportService.getOrderedExcelColumnEntity(Product.class);

        excelRowNames.forEach(entity -> {
            String columnName = entity.getColumnName();
            if ("id".equals(columnName)) {
                assertTrue(entity.getIndex().equals(10));
            }
            if ("价格".equals(columnName)) {
                assertTrue(entity.getIndex().equals(20));
            }
            if ("名称".equals(columnName)) {
                assertTrue(entity.getIndex().equals(40));
            }
            if ("订单状态".equals(columnName)) {
                assertTrue(entity.getIndex().equals(50));
                assertTrue(entity.getEnumKey().equals("code"));
            }
            if ("创建日期".equals(columnName)) {
                assertTrue(entity.getIndex().equals(60));
            }
        });

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
    public void testSheetHeaderIsEmpty() {
        assertTrue(StringUtils.isEmpty(ExcelExporter.create(ExcelType.XLSX).getSheetHeader()));
    }

    @Test
    public void testSheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelExporter excelExporter = ExcelExporter.create(ExcelType.XLSX).sheetheader(sheetHeader);
        assertTrue(excelExporter.getSheetHeader().equals(sheetHeader));
    }

}