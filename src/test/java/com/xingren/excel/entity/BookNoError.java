package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import lombok.Data;

/**
 * @author guang
 * @since 2020/3/17 3:38 下午
 */
@Data
public class BookNoError {

    @ExcelColumn(index = 10, columnName = "书名", required = true)
    private String bookName;

    @ExcelColumn(index = 20, columnName = "作者", required = true)
    private String author;

    @ExcelColumn(index = 30, columnName = "第二作者")
    private String secondAuthor;

    @ExcelColumn(index = 40, columnName = "上架状态", enumKey = "name", required = true)
    private BookEnum bookStatue;
}
