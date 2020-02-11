package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.util.List;

/**
 * @author guang
 * @since 2020/2/10 9:46 上午
 */
public class ExcelReaderTest {

    String productFile_Xlsx;
    String productFile_Xls;

    @Before
    public void before() throws FileNotFoundException {

        // 获取桌面路径
        FileSystemView fsv = FileSystemView.getFileSystemView();
        String user = fsv.getHomeDirectory().getPath();
        productFile_Xlsx = user + "/product.xlsx";
        productFile_Xls = user + "/product.xls";

    }

    @Test
    public void testImportXls() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(productFile_Xls);
        List<Product> products = ExcelReader
                .read(fileInputStream, ExcelType.XLS)
                .startRowNum(1)
                .to(Product.class);
        Assert.assertNotNull(products);
    }

    @Test
    public void testImportXlsx() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(productFile_Xls);
        List<Product> products = ExcelReader
                .read(fileInputStream)
                .startRowNum(1)
                .to(Product.class);
        Assert.assertNotNull(products);
    }

}
