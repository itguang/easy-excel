package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;

/**
 * @author guang
 * @since 2020/3/1 9:31 上午
 */
@Getter
@Setter
@AllArgsConstructor
@NoArgsConstructor
public class Person {

    @ExcelColumn(columnName = "姓名", index = 10)
    private String name;

    @ExcelColumn(columnName = "性别", index = 20)
    private Gender gender;
}
