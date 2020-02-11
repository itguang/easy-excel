package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.export.ExcelService;
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

public class ExcelWriterTest {

    ArrayList<Product> products;
    String testResourcesPath;
    String productFile;

    @Before
    public void before() throws FileNotFoundException {

        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String user = fsv.getHomeDirectory().getPath();
        productFile = user + "/product.xls";

        String testResourcesPath = ExcelWriterTest.class.getResource("/").getPath();
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

        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .write(products, Product.class);

        File file = new File(productFile);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    public void testGetExcelRowNames() {
        ExcelService excelService = ExcelService.forClass(Product.class);
        List<ExcelColumnAnnoEntity> excelRowNames = excelService.getOrderedExcelColumnEntity();

        excelRowNames.forEach(entity -> {
            String columnName = entity.getColumnName();
            if ("id".equals(columnName)) {
                assertTrue(entity.getIndex().equals(10));
            }
            if ("价格".equals(columnName)) {
                assertTrue(entity.getIndex().equals(20));
                assertTrue(entity.getSuffix().equals(" 元"));
                assertTrue(entity.getCentToYuan().equals(true));
            }
            if ("名称".equals(columnName)) {
                assertTrue(entity.getIndex().equals(40));
            }
            if ("订单状态".equals(columnName)) {
                assertTrue(entity.getIndex().equals(50));
                assertTrue(entity.getEnumKey().equals("name"));
                assertTrue(entity.getPrefix().equals("状态: "));
            }
            if ("创建日期".equals(columnName)) {
                assertTrue(entity.getIndex().equals(60));
            }
            if ("是否是新品".equals(columnName)) {
                assertTrue(entity.getIndex().equals(41));
                assertTrue(entity.getTrueStr().equals("新品"));
                assertTrue(entity.getFalseStr().equals("非新品"));
            }
        });

    }

    @Test
    public void testSheetName() {
        String sheetName = "测试 sheet name";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX)
                .sheetName(sheetName);
        assertTrue(excelWriter.getSheetName().equals(sheetName));
    }

    @Test
    public void testDefaultSheetName() {
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX);
        assertTrue(ExcelConstant.DEFAULT_SHEET_NAME.equals(excelWriter.getSheetName()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void testSheetNameIsEmpty() {
        ExcelWriter.create(ExcelType.XLSX).sheetName(null);
    }

    @Test
    public void testSheetHeaderIsEmpty() {
        assertTrue(StringUtils.isEmpty(ExcelWriter.create(ExcelType.XLSX).getSheetHeader()));
    }

    @Test
    public void testSheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX).sheetHeader(sheetHeader);
        assertTrue(excelWriter.getSheetHeader().equals(sheetHeader));
    }

}