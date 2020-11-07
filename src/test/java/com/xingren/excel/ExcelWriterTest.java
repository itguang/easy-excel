package com.xingren.excel;

import com.xingren.excel.entity.*;
import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.pojo.TemplateRowFormat;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.util.DateUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.*;
import java.net.URLDecoder;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static com.xingren.excel.enums.FormatType.DROP_DOWN_LIST;
import static java.util.stream.Collectors.toList;
import static org.junit.jupiter.api.Assertions.*;

public class ExcelWriterTest {

    /*
        导出文件在: out/test/resources/export 目录下
     */

    static ArrayList<Product> products;
    static String productFile_XLS;
    static String productFile_with_tail_XLS;
    static String productNoHeaderFile_XLS;
    static String productFile_XLSX;
    static String simpleProductFile_XLSX;
    static String productEmptyFile_XLSX;
    static String employeeTemplate_XLSX;
    static String bookTemplate_XLSX;
    static String orderTemplate_XLSX;
    static String orderWithErrorInfo_XLSX;
    static String orderNoErrorInfo_XLSX;
    static String multi_sheet_template_XLSX;
    static String multi_sheet_XLSX;
    static String format_employ_sheet_XLSX;

    static ArrayList<SimpleProduct> simpleProducts;
    static ArrayList<Product> products_1w;

    public ExcelWriterTest() throws UnsupportedEncodingException {
        ClassLoader classLoader = ExcelWriterTest.class.getClassLoader();
        productFile_XLS = URLDecoder.decode(classLoader.getResource("export/导出商品数据.xls").getPath(), "utf-8");
        productFile_with_tail_XLS = URLDecoder.decode(
                classLoader.getResource("export/导出商品数据(tail).xls").getPath(), "utf-8");
        productNoHeaderFile_XLS = URLDecoder.decode(
                classLoader.getResource("export/no_header.xls").getPath(), "utf-8");
        productFile_XLSX = URLDecoder.decode(classLoader.getResource("export/导出商品数据_1w.xlsx").getPath(),
                "utf-8");
        simpleProductFile_XLSX = URLDecoder.decode(
                classLoader.getResource("export/导出商品简单数据_1w.xlsx").getPath(), "utf-8");
        productEmptyFile_XLSX = URLDecoder.decode(classLoader.getResource("export/空的商品数据.xlsx").getPath(),
                "utf-8");
        employeeTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/员工模板表.xlsx").getPath(),
                "utf-8");
        bookTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/图书模板表.xlsx").getPath(),
                "utf-8");
        orderTemplate_XLSX = URLDecoder.decode(classLoader.getResource("export/订单模板表.xlsx").getPath(), "utf-8");
        orderWithErrorInfo_XLSX = URLDecoder.decode(classLoader.getResource("export/订单错误信息.xlsx").getPath(), "utf-8");
        orderNoErrorInfo_XLSX = URLDecoder.decode(classLoader.getResource("export/订单没有错误信息.xlsx").getPath(), "utf-8");
        multi_sheet_template_XLSX = URLDecoder.decode(classLoader.getResource("export/多sheet模板导出.xlsx").getPath(),
                "utf-8");
        multi_sheet_XLSX = URLDecoder.decode(classLoader.getResource("export/多sheet导出.xlsx").getPath(), "utf-8");
        format_employ_sheet_XLSX = URLDecoder.decode(classLoader.getResource("export/格式化的员工表.xlsx").getPath(), "utf-8");

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

        simpleProducts = new ArrayList<>();
        for (int i = 0; i < 10000; i++) {
            SimpleProduct simpleProduct = new SimpleProduct(i,
                    1000L, "2020-01-01",
                    "苹果" + i, "是", StateEnum.DOWN.getName(), "2020-08-08",
                    "好吃" + i, "");
            simpleProducts.add(simpleProduct);
        }

        products_1w = new ArrayList<>();
        for (int i = 0; i < 10000; i++) {
            Product pp = new Product(i,
                    1000L, OffsetDateTime.now(),
                    "苹果" + i, true, StateEnum.DOWN, LocalDateTime.now(),
                    "好吃" + i);
            products_1w.add(pp);
        }

    }

    @Test
    @DisplayName("导出商品数据(测试数据正确性)")
    public void test_export_xls() throws IOException {
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .write(products, Product.class);
        System.out.println("=============> productFile_XLS = " + productFile_XLS);
        File file = new File(productFile_XLS);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        // 读取该文件断言数据是否符合预期

        FileInputStream inputStream = new FileInputStream(file);
        List<Product> products = ExcelReader.read(inputStream, ExcelType.XLS)
                .columnNameRowNum(1)
                .toPojo(Product.class);
        inputStream.close();

        Assertions.assertEquals(2, products.size());
        Product product = products.get(0);

        assertEquals(1000, (int) product.getId());
        assertEquals(1000, (long) product.getPrice());

        assertEquals(StateEnum.DOWN, product.getStateEnum());
        assertTrue(product.getIsNew());

        assertEquals("苹果--好吃", product.getOther());

        assertEquals(LocalDateTime.of(LocalDate.now(), LocalTime.MIN), product.getUpdateTime());
        String time = DateUtil.formatOffsetDateTime(OffsetDateTime.now(), "yyyy-MM-dd");
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime(time, "yyyy-MM-dd");
        assertEquals(offsetDateTime, product.getCreated());

    }

    @Test
    @DisplayName("导出商品数据(tail)")
    public void test_export_with_tail_xls() throws IOException {
        String sheetRemark = "附: 统计说明\n"
                + "1. 新增银企鹅比例=新增银企鹅数/(就诊总人数-(银企鹅总数-新增银企鹅数)),分母即为: 该时间段内就诊的并且在该时间段之前不是银企鹅的总人数;\n"
                + "2. 就诊人数以接口传输至企鹅侧的数据为准. 就诊人数(开始就诊时间计)为: 所选择的时间范围内,就诊的人数(去重,即一个人就诊2次算1);\n"
                + "3. 银企鹅数量: 该时间段内就诊人中银企鹅会员的人数;\n"
                + "4. 新增银企鹅数: 该时间段内成为银企鹅的人数.";
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .sheetTail(sheetRemark, 22f)
                .write(products, Product.class);

        File file = new File(productFile_with_tail_XLS);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    @DisplayName("导出没有 header 的 xls")
    public void test_export_no_sheet_header_xls() throws IOException {
        Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                .sheetName("商品数据")
                .write(products, Product.class);

        File file = new File(productNoHeaderFile_XLS);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        FileInputStream inputStream = new FileInputStream(file);
        List<Product> products = ExcelReader.read(inputStream, ExcelType.XLS)
                .columnNameRowNum(0)
                .toPojo(Product.class);
        inputStream.close();

        Assertions.assertEquals(2, products.size());
        Product product = products.get(1);

        assertEquals(2000, (int) product.getId());
        assertEquals(2000, (long) product.getPrice());

        assertEquals(StateEnum.UP, product.getStateEnum());
        assertFalse(product.getIsNew());

        assertEquals("橘子--好酸", product.getOther());

        assertEquals(LocalDateTime.of(LocalDate.now().minusDays(1), LocalTime.MIN), product.getUpdateTime());
        String time = DateUtil.formatOffsetDateTime(OffsetDateTime.now().minusDays(1), "yyyy-MM-dd");
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime(time, "yyyy-MM-dd");
        assertEquals(offsetDateTime, product.getCreated());

    }

    @Test
    @DisplayName("导出带有 header 和 columnName 的空模板")
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

        FileInputStream fileInputStream = new FileInputStream(file);

        List<Employee> employees = ExcelReader.read(fileInputStream)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        fileInputStream.close();

        assertEquals(0, employees.size());

    }

    @Test
    @DisplayName("测试没有指定 POJO 中的枚举 值异常")
    public void test_export_no_enumKey() {
        Person person = new Person("小李", Gender.MALE);
        ArrayList<Person> persons = new ArrayList<>();
        persons.add(person);

        try {
            Workbook workbook = ExcelWriter.create(ExcelType.XLS)
                    .sheetName("人员数据")
                    .write(persons, Person.class);
        } catch (Exception e) {
            assertEquals("类 com.xingren.excel.entity.Person 中枚举字段:gender 没有指定 enumKey 值!", e.getMessage());
        }
    }

    @Test
    @DisplayName("导出空的商品数据")
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

        FileInputStream fileInputStream = new FileInputStream(file);
        List<Product> products = ExcelReader.read(fileInputStream)
                .columnNameRowNum(1)
                .toPojo(Product.class);

        assertEquals(0, products.size());
        fileInputStream.close();

    }

    /**
     * 10000 行数据导出测试 7s 需要优化
     */
    @Test
    @DisplayName(" 1w 行数据导出测试")
    public void test_export_1w_xlsx() throws IOException {

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .activeSheet(0)
                .write(products_1w, Product.class);

        File file = new File(productFile_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        FileInputStream fileInputStream = new FileInputStream(file);
        List<Product> list = ExcelReader.read(fileInputStream)
                .columnNameRowNum(1)
                .toPojo(Product.class);
        assertEquals(10000, list.size());

        fileInputStream.close();

    }

    /**
     * 10000 行数据导出测试 7s 需要优化
     */
    @Test
    @DisplayName(" 1w 行简单类型数据导出测试")
    public void test_export_1w_simple_xlsx() throws IOException {

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("商品数据")
                .sheetHeader("--2月份商品数据--")
                .activeSheet(0)
                .write(simpleProducts, SimpleProduct.class);

        File file = new File(simpleProductFile_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        FileInputStream fileInputStream = new FileInputStream(file);
        List<SimpleProduct> list = ExcelReader.read(fileInputStream)
                .columnNameRowNum(1)
                .toPojo(SimpleProduct.class);
        assertEquals(10000, list.size());

        fileInputStream.close();

    }

    @Test
    @DisplayName("测试 注解解析 ExcelColumnAnnoEntity 的值")
    public void test_get_excel_rowNames() {
        List<ExcelColumnAnnoEntity> excelRowNames =
                ExcelColumnService.forClass(Product.class).getOrderedExcelColumnEntity();

        excelRowNames.forEach(entity -> {
            String columnName = entity.getColumnName();
            if ("id".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 10);
            }
            if ("价格".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 20);
                assertEquals(entity.getSuffix(), " 元");
                assertEquals(entity.getCentToYuan(), true);
            }
            if ("名称".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 40);
            }
            if ("订单状态".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 50);
                assertEquals(entity.getEnumKey(), "name");
                assertEquals(entity.getPrefix(), "状态: ");
            }
            if ("创建日期".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 60);
            }
            if ("是否是新品".equals(columnName)) {
                assertEquals((int) entity.getIndex(), 41);
                assertEquals(entity.getTrueStr(), "新品");
                assertEquals(entity.getFalseStr(), "非新品");
            }
        });

    }

    @Test
    @DisplayName("获取 sheet name ")
    public void test_sheetName() {
        String sheetName = "测试 sheet name";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX)
                .sheetName(sheetName);
        assertEquals(sheetName, excelWriter.getSheetName());
    }

    @Test
    @DisplayName("测试获取默认的 sheet name")
    public void test_default_sheetName() {
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX);
        assertEquals(excelWriter.getSheetName(), ExcelConstant.DEFAULT_SHEET_NAME);
    }

    @Test
    @DisplayName("测试 sheet name 不能为空")
    public void test_sheetName_is_empty() {
        try {
            ExcelWriter.create(ExcelType.XLSX).sheetName(null);
        } catch (Exception e) {
            assertEquals("sheetName cannot be empty", e.getMessage());
        }
    }

    @Test
    @DisplayName("创建空的 sheet header ")
    public void test_sheetHeader_is_empty() {
        assertTrue(StringUtils.isEmpty(ExcelWriter.create(ExcelType.XLSX).getSheetHeader()));
    }

    @Test
    @DisplayName("测试 sheet header 设置")
    public void test_sheetHeader() {
        String sheetHeader = "测试 sheet header";
        ExcelWriter excelWriter = ExcelWriter.create(ExcelType.XLSX).sheetHeader(sheetHeader);
        assertEquals(sheetHeader, excelWriter.getSheetHeader());
    }

    @Test
    @DisplayName("测试继承了 ErrorInfoRow 的模板导出")
    public void test_write_order_template() throws Exception {
        String sheetHeader = "继承了 ErrorInfoRow 的模板导出";
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader).writeTemplate(OrderEntity.class);

        File file = new File(orderTemplate_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    @DisplayName("测试有错误消息的导出")
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
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader)
                .write(orderEntities, OrderEntity.class);

        File file = new File(orderWithErrorInfo_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    @DisplayName("测试没有错误信息但继承了 ErrorInfoRow 的导出")
    public void test_write_no_error_order() throws Exception {

        ArrayList<OrderEntity> orderEntities = new ArrayList<>();
        orderEntities.add(new OrderEntity(1, "45674567890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));
        orderEntities.add(new OrderEntity(2, "345678907890-jk"));

        String sheetHeader = "继承了 ErrorInfoRow 的导出";
        Workbook workbook = ExcelWriter.create().sheetHeader(sheetHeader)
                .write(orderEntities, OrderEntity.class);

        File file = new File(orderNoErrorInfo_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();
    }

    @Test
    @DisplayName("测试导出图书模板")
    public void test_write_template_book() throws Exception {

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("图书模板表").writeTemplate(Book.class);

        File file = new File(bookTemplate_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

    }

    @Test
    @DisplayName("测试多 sheet 模板导出")
    public void write_multi_sheet_template() throws Exception {

        Workbook workbook = ExcelWriter.create().addSheetTemplate("图书模板", "图书", Book.class)
                .addSheetTemplate("订单模板", OrderEntity.class)
                .getWorkbook();

        File file = new File(multi_sheet_template_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        // 多 sheet 读取
        FileInputStream fileInputStream = new FileInputStream(file);
        List<Book> books = ExcelReader.read(fileInputStream).toPojo(0, 1, Book.class);

    }

    @Test
    @DisplayName("测试多 sheet 带数据的导出")
    public void write_multi_sheet() throws Exception {

        ArrayList<OrderEntity> orders = new ArrayList<>();
        OrderEntity orderEntity = new OrderEntity(100, "fghjkl");
        orders.add(orderEntity);

        Workbook workbook = ExcelWriter.create().addSheet("第一个", "商品数据", products, Product.class)
                .addSheet("第二个", orders, OrderEntity.class)
                .getWorkbook();

        File file = new File(multi_sheet_XLSX);
        OutputStream outputStream = new FileOutputStream(file);
        workbook.write(outputStream);
        outputStream.close();

        // 多 sheet 读取
        InputStream fileInputStream = new FileInputStream(new File(multi_sheet_XLSX));

        ExcelReader read = ExcelReader.read(fileInputStream);
        List<Product> readProductList = read
                .toPojo(0, 1, Product.class);
        assertEquals(2, readProductList.size());
        assertEquals(readProductList.get(0).getId(), products.get(0).getId());
        assertEquals(readProductList.get(1).getName(), products.get(1).getName());

        List<OrderEntity> orderEntities = read.toPojo(1, 0, OrderEntity.class);

        assertEquals(1, orderEntities.size());
        assertEquals(orderEntity, orderEntities.get(0));
    }

    @Test
    @DisplayName("测试多sheet 导出, sheetName 不能为空")
    public void write_multi_sheet_with_empty_sheet_name() throws Exception {

        ArrayList<OrderEntity> orders = new ArrayList<>();
        orders.add(new OrderEntity(100, "fghjkl"));

        Workbook workbook = null;
        try {
            workbook = ExcelWriter.create().addSheet("", "商品数据", products, Product.class)
                    .addSheet("第二个", orders, OrderEntity.class)
                    .getWorkbook();
            File file = new File(multi_sheet_XLSX);
            OutputStream outputStream = new FileOutputStream(file);
            workbook.write(outputStream);
        } catch (Exception e) {
            assertEquals("第 1 个 sheetName 不能为空", e.getMessage());
        }

    }

    @Test
    @DisplayName("测试格式化的导出")
    public void test_export_format_template() throws IOException {
        List<TemplateRowFormat> formats = new ArrayList<>();
        TemplateRowFormat format = new TemplateRowFormat(2, 4, 2, 4);
        format.setFormatType(DROP_DOWN_LIST);
        List<String> genders = Stream.of(Gender.values()).map(Gender::getName).collect(toList());
        String[] genderArr = new String[genders.size()];
        genders.toArray(genderArr);
        format.setData(genderArr);
        formats.add(format);

        Employee employee = new Employee();
        employee.setAge(18);
        employee.setGender(Gender.FEMALE);
        employee.setName("张三");
        ArrayList<Employee> list = new ArrayList<>();
        list.add(employee);

        Workbook workbook = ExcelWriter.create(ExcelType.XLSX)
                .sheetName("格式化的员工表")
                .sheetHeader("--员工数据--")
                .activeSheet(0)
                .write(list, Employee.class, formats);

        OutputStream outputStream = new FileOutputStream(format_employ_sheet_XLSX);
        workbook.write(outputStream);
        outputStream.close();

    }
}