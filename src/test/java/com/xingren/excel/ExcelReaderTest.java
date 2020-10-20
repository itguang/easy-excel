package com.xingren.excel;

import com.xingren.excel.entity.*;
import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.util.DateUtil;
import org.apache.commons.lang3.StringUtils;
import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.time.OffsetDateTime;
import java.util.List;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;

/**
 * @author guang
 * @since 2020/2/10 9:46 上午
 */
public class ExcelReaderTest {

    private static String employeeFile_Xlsx;
    private static String bookFile_Xlsx;
    private static String bookEnumFile_Xlsx;
    private static String employeeFile_Xls;
    private static String employeeEmptyFile_Xls;
    private static String employeeFile_1w_Xlsx;
    private static String multi_sheet_Xlsx;

    public ExcelReaderTest() throws UnsupportedEncodingException {
        ClassLoader classLoader = ExcelReaderTest.class.getClassLoader();
        employeeFile_Xls = URLDecoder.decode(classLoader.getResource("read/员工数据.xls").getPath(), "utf-8");
        bookFile_Xlsx = URLDecoder.decode(classLoader.getResource("read/图书数据.xlsx").getPath(), "utf-8");
        bookEnumFile_Xlsx = URLDecoder.decode(classLoader.getResource("read/图书数据-枚举值为空.xlsx").getPath(), "utf-8");
        employeeEmptyFile_Xls = URLDecoder.decode(classLoader.getResource("read/空员工数据.xls").getPath(), "utf-8");
        employeeFile_Xlsx = URLDecoder.decode(classLoader.getResource("read/员工数据.xlsx").getPath(), "utf-8");
        employeeFile_1w_Xlsx = URLDecoder.decode(classLoader.getResource("read/员工数据_1w.xlsx").getPath(), "utf-8");
        multi_sheet_Xlsx = URLDecoder.decode(classLoader.getResource("read/多sheet导入.xlsx").getPath(), "utf-8");
    }

    @Test
    @DisplayName("读取一个空 xls")
    public void test_read_empty() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeEmptyFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assertions.assertEquals(0, employees.size());
    }

    /**
     * 1w 行数据导入测试 -- 3s
     */
    @Test
    @DisplayName("1w 行数据导入")
    public void test_import_1w_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_1w_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assertions.assertEquals(10000, employees.size());
    }

    @Test
    @DisplayName("测试读取数据正确性")
    public void test_import_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assertions.assertEquals(2, employees.size());
        Employee employee = employees.get(1);
        Assertions.assertEquals(20000L, (long) employee.getId());
        Assertions.assertEquals(25, (int) employee.getAge());
        // MyReadConverter.class
        Assertions.assertEquals("[李红]", employee.getName());
        // 去掉后缀: 元, 并且: 元转分
        Assertions.assertEquals(12000000L, (long) employee.getSalary());
        // 枚举转换
        Assertions.assertEquals(Gender.FEMALE, employee.getGender());
        // 日期处理
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime("2020/2/25", DEFAULT_DATE_PATTREN);
        Assertions.assertEquals(offsetDateTime, employee.getBirthday());

        // Boolean 值
        Assertions.assertFalse(employee.getIsDown());

    }

    /**
     * 起始行不合法
     */
    @Test
    @DisplayName("起始行不合法")
    public void test_exception_columnNameRowNum_import_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xls);
        try {
            ExcelReader
                    .read(fileInputStream, ExcelType.XLS)
                    .columnNameRowNum(0)
                    .toPojo(Employee.class);
        } catch (Exception e) {
            Assertions.assertEquals("解析 Excel 失败,检查 columnNameRowNum 是否设置正确", e.getMessage());
        }
    }

    @Test
    @DisplayName("读取为部分为空的 excel cell 值")
    public void test_import_xlsx() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .columnNameRowNum(1)
                .toPojo(Employee.class);

        Assertions.assertEquals(3, employees.size());

        Employee employee = employees.get(2);

        Assertions.assertEquals(3, (long) employee.getId());
        Assertions.assertEquals("[]", employee.getName());
        Assertions.assertNull(employee.getSalary());
        Assertions.assertNull(employee.getBirthday());
        Assertions.assertNull(employee.getIsDown());
        Assertions.assertEquals(Gender.MALE, employee.getGender());
    }

    @Test
    @DisplayName("测试 require 和 ErrorInfo 功能")
    public void test_import_requied() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(bookFile_Xlsx);
        List<Book> books = ExcelReader
                .read(fileInputStream)
                .toPojo(Book.class);

        Assertions.assertTrue(StringUtils.isEmpty(books.get(0).getErrorInfo()));
        Assertions.assertEquals(books.get(1).getErrorInfo(), "[作者]不能为空\n");
    }

    /**
     * com.xingren.excel.exception.ExcelException: 枚举值[不存在的枚举值]不合法
     */
    @Test
    @DisplayName("测试 require, 枚举值映射 和 ErrorInfo 功能")
    public void test_import_enum_value_is_null_error_row() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(bookEnumFile_Xlsx);
        List<Book> books = ExcelReader.read(fileInputStream).toPojo(Book.class);

        Assertions.assertEquals("[上架状态]不合法\n", books.get(2).getErrorInfo());
        Assertions.assertEquals("[上架状态]不合法\n", books.get(1).getErrorInfo());
        Assertions.assertNull(books.get(2).getBookStatue());
        Assertions.assertNull(books.get(1).getBookStatue());

        Assertions.assertEquals(BookEnum.UP, books.get(0).getBookStatue());
        Assertions.assertEquals("", books.get(0).getErrorInfo());

    }

    @Test
    @DisplayName("测试 枚举值映射 的运行时异常")
    public void test_import_enum_value_is_null() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(bookEnumFile_Xlsx);
        try {
            List<BookNoError> books = ExcelReader
                    .read(fileInputStream)
                    .toPojo(BookNoError.class);
        } catch (Exception e) {
            Assertions.assertEquals("[上架状态]不合法", e.getMessage());
        }
    }

    @Test
    @DisplayName("测试 多 sheet 读取")
    public void test_mult_sheet() throws FileNotFoundException {
        InputStream fis = new FileInputStream(multi_sheet_Xlsx);
        ExcelReader reader = ExcelReader.read(fis);
        List<Book> books = reader.toPojo(0, 0, Book.class);
        List<Book> empty = reader.toPojo(1, 0, Book.class);
        List<Employee> employees = reader.toPojo(2, 1, Employee.class);
        Assertions.assertEquals(books.size(), 2);

        Assertions.assertEquals(empty.size(), 0);
        Assertions.assertEquals(employees.size(), 2);

        Employee employee = employees.get(0);
        Assertions.assertEquals(10000L, (long) employee.getId());
        Assertions.assertEquals("[张三]", employee.getName());
        Assertions.assertEquals(20, (int) employee.getAge());
        Assertions.assertEquals(10000000, (long) employee.getSalary());
        Assertions.assertEquals(Gender.MALE, employee.getGender());
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime("1994/4/27", DEFAULT_DATE_PATTREN);
        Assertions.assertEquals(offsetDateTime, employee.getBirthday());

        // Boolean 值
        Assertions.assertTrue(employee.getIsDown());

    }
}
