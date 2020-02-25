package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import lombok.Data;

import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/2/25 10:47 上午
 */
@Data
public class Employee {

    @ExcelColumn(index = 10, columnName = "ID")
    private Integer id;
    @ExcelColumn(index = 20, columnName = "名称")
    private String name;
    @ExcelColumn(index = 30, columnName = "工资", centToYuan = true)
    private Long salary;
    @ExcelColumn(index = 40, columnName = "性别")
    private Gender gender;
    @ExcelColumn(index = 50, columnName = "生日",datePattern = "yyyy-MM-dd")
    private OffsetDateTime birthday;
}
