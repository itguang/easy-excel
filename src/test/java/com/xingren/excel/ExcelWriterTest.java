package com.xingren.excel;

import com.xingren.excel.entity.*;
import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.service.ExcelColumnService;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.*;
import java.net.URLDecoder;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.List;

import static org.junit.Assert.assertTrue;

public class ExcelWriterTest {

    static ArrayList<Product> products;
    static String productFile_XLS;
    static String productNoHeaderFile_XLS;
    static String productFile_XLSX;
    static String productEmptyFile_XLSX;
    static String employeeTemplate_XLSX;
    static String bookTemplate_XLSX;
    static String orderTemplate_XLSX;
    static String orderWithErrorInfo_XLSX;
    static String orderNoErrorInfo_XLSX;

    @BeforeClass
    public static void before() throws UnsupportedEncodingException {
        ClassLoader classLoader = ExcelWriterTest.class.getClassLoader();
        productFile_XLS = URLDecoder.decode(classLoader.getResource("export/导出商品数据.xls").getPath(), "utf-8");
        productNoHeaderFile_XLS = URLDecoder.decode(classLoader.getResource("export/no_header.xls").getPath(), "utf-8");
        productFile_XLSX = URLDecoder.decode(classLoader.getResource("export/导出商品数据_1w.xlsx").getPath(), "utf-8");
        productEmptyFile_XLSX = URLDecoder.decode(classLoader.getResource("export/空的商品数据.xlsx").getPath(), "utf-8");
        employeeTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/员工模板表.xlsx").getPath(), "utf-8");
        bookTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/图书模板表.xlsx").getPath(), "utf-8");
        orderTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/订单模板表.xlsx").getPath(), "utf-8");
        orderWithErrorInfo_XLSX = URLDecoder.decode(classLoader.getResource("export/订单错误信息.xlsx").getPath(), "utf-8");
        orderNoErrorInfo_XLSX = URLDecoder.decode(classLoader.getResource("export/订单没有错误信息.xlsx").getPath(), "utf-8");

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
    public void test_export_xls() throws IOException {
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .write(products, Product.class);

        File file = new File(productFile_XLS);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    public void test_export_no_sheet_header_xls() throws IOException {
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .write(products, Product.class);

        File file = new File(productNoHeaderFile_XLS);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    public void test_export_employee_template() throws IOException {
        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("员工模板表")
                .sheetHeader("--员工数据--")
                .activeSheet(0)
                .writeTemplate(Employee.class);

        File file = new File(employeeTemplate_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test(expected = ExcelConvertException.class)
    public void test_export_no_enumKey() {
        Person person = new Person("小李", Gender.MALE);
        ArrayList<Person> persons = new ArrayList<>();
        persons.add(person);
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("人员数据")
                .write(persons, Person.class);
    }

    @Test
    public void test_export_empty_xlsx() throws IOException {
        ArrayList<Product> emptyProducts = new ArrayList<>();
        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("空的商品数据")
                .sheetHeader("--2月份商品数据--")
                .activeSheet(0)
                .write(emptyProducts, Product.class);

        File file = new File(productEmptyFile_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    /**
     * 10000 行数据导出测试 7s 需要优化
     */
    @Test
    public void test_export_1w_xlsx() throws IOException {

        ArrayList<Product> products = new ArrayList<>();

        for (int i = 0; i < 10000; i++) {
            Product apple = new Product(i,
                    1000L, OffsetDateTime.now(),
                    "苹果" + i, true, StateEnum.DOWN, LocalDateTime.now(),
                    "好吃" + i);
            products.add(apple);
        }

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .activeSheet(0)
                .write(products, Product.class);

        File file = new File(productFile_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    public void test_get_excel_rowNames() {
        List<ExcelColumnAnnoEntity> excelRowNames =
                ExcelColumnService.forClass(Product.class).getOrderedExcelColumnEntity();

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
    public void test_sheetName() {
        String sheetName = "测试 sheet name";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX)
                .sheetName(sheetName);
        assertTrue(excelWriter.getSheetName().equals(sheetName));
    }

    @Test
    public void test_default_sheetName() {
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX);
        assertTrue(ExcelConstant.DEFAULT_SHEET_NAME.equals(excelWriter.getSheetName()));
    }

    @Test(expected = IllegalArgumentException.class)
    public void test_sheetName_is_empty() {
        ExcelWriter.create(ExcelType.XLSX).sheetName(null);
    }

    @Test
    public void test_sheetHeader_is_empty() {
        assertTrue(StringUtils.isEmpty(ExcelWriter.create(ExcelType.XLSX).getSheetHeader()));
    }

    @Test
    public void test_sheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX).sheetHeader(sheetHeader);
        assertTrue(excelWriter.getSheetHeader().equals(sheetHeader));
    }

    @Test
    public void test_write_order_template() throws Exception {
        String sheetHeader = "继承了 ErrorInfoRow 的模板导出";
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader).writeTemplate(OrderEntity.class);

        File file = new File(orderTemplate_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    public void test_write_order_with_error_info() throws Exception {

        ArrayList<OrderEntity> orderEntities = new ArrayList<>();
        orderEntities.add(new OrderEntity(1, "45674567890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));

        orderEntities.forEach(orderEntity -> {
            StringBuilder sb = new StringBuilder();
            sb.append("订单号[" + orderEntity.getOrderNo() + "]不对! \n");
            orderEntity.setErrorInfo(sb.toString());
        });

        String sheetHeader = "继承了 ErrorInfoRow 的导出";
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader).write(orderEntities, OrderEntity.class);

        File file = new File(orderWithErrorInfo_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    public void test_write_no_error_order() throws Exception {

        ArrayList<OrderEntity> orderEntities = new ArrayList<>();
        orderEntities.add(new OrderEntity(1, "45674567890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));

        String sheetHeader = "继承了 ErrorInfoRow 的导出";
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader).write(orderEntities, OrderEntity.class);

        File file = new File(orderNoErrorInfo_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    public void test_write_template_book() throws Exception {

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("图书模板表").writeTemplate(Book.class);

        File file = new File(bookTemplate_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

}