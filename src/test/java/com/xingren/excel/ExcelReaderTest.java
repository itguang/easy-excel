package com.xingren.excel;

import com.xingren.excel.entity.Employee;
import com.xingren.excel.enums.ExcelType;
import com.xingren.excel.exception.ExcelException;
import org.junit.Assert;
import org.junit.BeforeClass;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.io.UnsupportedEncodingException;
import java.net.URLDecoder;
import java.util.List;

/**
 * @author guang
 * @since 2020/2/10 9:46 上午
 */
public class ExcelReaderTest {

    private static String employeeFile_Xlsx;
    private static String employeeFile_Xls;
    private static String employeeEmptyFile_Xls;
    private static String employeeFile_1w_Xlsx;

    @BeforeClass
    public static void before() throws UnsupportedEncodingException {
        ClassLoader classLoader = ExcelReaderTest.class.getClassLoader();
        employeeFile_Xls = URLDecoder.decode(classLoader.getResource("read/员工数据.xls").getPath(), "utf-8");
        employeeEmptyFile_Xls = URLDecoder.decode(classLoader.getResource("read/空员工数据.xls").getPath(), "utf-8");
        employeeFile_Xlsx = URLDecoder.decode(classLoader.getResource("read/员工数据.xlsx").getPath(), "utf-8");
        employeeFile_1w_Xlsx = URLDecoder.decode(classLoader.getResource("read/员工数据_1w.xlsx").getPath(), "utf-8");

    }

    @Test
    public void test_read_empty() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeEmptyFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assert.assertNotNull(employees.size() == 0);
    }

    /**
     * 1w 行数据导入测试 -- 3s
     *
     * @throws FileNotFoundException
     */
    @Test
    public void test_import_1w_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_1w_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assert.assertNotNull(employees);
    }

    @Test
    public void test_import_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assert.assertNotNull(employees);
    }

    /**
     * 起始行不合法
     *
     * @throws FileNotFoundException
     */
    @Test(expected = ExcelException.class)
    public void test_exception_columnNameRowNum_import_xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .columnNameRowNum(0)
                .toPojo(Employee.class);
        Assert.assertNotNull(employees);
    }

    @Test
    public void test_import_xlsx() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .columnNameRowNum(1)
                .toPojo(Employee.class);
        Assert.assertNotNull(employees);
    }

}
