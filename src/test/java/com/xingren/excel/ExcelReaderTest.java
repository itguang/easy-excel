package com.xingren.excel;

import com.xingren.excel.entity.Employee;
import com.xingren.excel.enums.ExcelType;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.InputStream;
import java.util.List;

/**
 * @author guang
 * @since 2020/2/10 9:46 上午
 */
public class ExcelReaderTest {

    private String employeeFile_Xlsx;
    private String employeeFile_Xls;
    private String employeeFile_1w_Xlsx;

    @Before
    public void before() {
        String resourcePath = this.getClass().getClassLoader().getResource("").getPath().replace("classes",
                "resources");
        employeeFile_Xls = resourcePath + "read/员工数据.xls";
        employeeFile_Xlsx = resourcePath + "read/员工数据.xlsx";
        employeeFile_1w_Xlsx = resourcePath + "read/员工数据_1w.xlsx";

    }

    /**
     * 1w 行数据导入测试 -- 3s
     *
     * @throws FileNotFoundException
     */
    @Test
    public void testImport_1w_Xls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_1w_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .startRowNum(1)
                .to(Employee.class);
        Assert.assertNotNull(employees);
    }

    @Test
    public void testImportXls() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xls);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .startRowNum(1)
                .to(Employee.class);
        Assert.assertNotNull(employees);
    }

    @Test
    public void testImportXlsx() throws FileNotFoundException {
        InputStream fileInputStream = new FileInputStream(employeeFile_Xlsx);
        List<Employee> employees = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .startRowNum(1)
                .to(Employee.class);
        Assert.assertNotNull(employees);
    }

}
