package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

/**
 * @author guang
 * @since 2020/1/18 10:32 上午
 */
@Data
@NoArgsConstructor
public class SimpleProduct {

    public SimpleProduct(Integer id, Long price, String created, String name, String isNew, String stateEnum,
                         String updateTime, String other, String test) {
        this.id = id;
        this.price = price;
        this.created = created;
        this.name = name;
        this.isNew = isNew;
        this.stateEnum = stateEnum;
        this.updateTime = updateTime;
        this.other = other;
        this.test = test;
    }

    @ExcelColumn(columnName = "id", index = 10)
    private Integer id;

    @ExcelColumn(columnName = "价格", index = 20)
    private Long price;

    @ExcelColumn(columnName = "创建日期", index = 60)
    private String created;

    @ExcelColumn(columnName = "名称", index = 40)
    private String name;

    @ExcelColumn(columnName = "是否是新品", index = 41)
    private String isNew;

    @ExcelColumn(columnName = "订单状态", index = 50)
    private String stateEnum;

    @ExcelColumn(columnName = "状态变更日期", index = 55)
    private String updateTime;

    @ExcelColumn(columnName = "备注", index = 70)
    private String other;

    private String test;

}
