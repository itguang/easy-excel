package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.MyReadConverter;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/2/25 10:47 上午
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Employee {

    @ExcelColumn(index = 10, columnName = "ID")
    private Long id;
    @ExcelColumn(index = 20, columnName = "名称", readConverter = MyReadConverter.class)
    private String name;
    @ExcelColumn(index = 15, columnName = "年龄")
    private Integer age;
    @ExcelColumn(index = 30, columnName = "工资", yuanToCent = true, suffix = " 元")
    private Long salary;
    @ExcelColumn(index = 40, columnName = "性别", enumKey = "name")
    private Gender gender;
    @ExcelColumn(index = 50, columnName = "生日")
    private OffsetDateTime birthday;
    @ExcelColumn(index = 60, columnName = "是否离职", strToTrue = "是", strToFalse = "否")
    private Boolean isDown;
}
