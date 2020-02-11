package com.xingren.excel;

import com.xingren.excel.enums.ExcelType;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.Assert;
import org.junit.Before;
import org.junit.Test;

import javax.swing.filechooser.FileSystemView;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
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
    public void poiImport_XlSX() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(productFile_Xlsx);
//        BufferedInputStream bufferedInputStream = new BufferedInputStream(fileInputStream);
//        POIFSFileSystem fileSystem = new POIFSFileSystem(bufferedInputStream);
        Workbook workbook = new XSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheet("商品数据");

        int lastRowIndex = sheet.getLastRowNum();

        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        String sheetHeader = cell.getStringCellValue();

        int firstRowNum = sheet.getFirstRowNum();

    }

    @Test
    public void poiImport_Xls() throws IOException {

        FileInputStream fileInputStream = new FileInputStream(productFile_Xls);
        Workbook workbook = new HSSFWorkbook(fileInputStream);

        Sheet sheet = workbook.getSheet("商品数据");

        int lastRowIndex = sheet.getLastRowNum();

        Row row = sheet.getRow(0);
        Cell cell = row.getCell(0);
        String sheetHeader = cell.getStringCellValue();

        int firstRowNum = sheet.getFirstRowNum();

    }

    @Test
    public void testImport() throws FileNotFoundException {
        FileInputStream fileInputStream = new FileInputStream(productFile_Xlsx);
        List<Product> products = ExcelReader
                .read(fileInputStream, ExcelType.XLSX)
                .startRowNum(1)
                .to(Product.class);
        Assert.assertNotNull(products);
    }

}
