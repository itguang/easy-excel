package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.MyConverter;
import com.xingren.excel.handler.impl.GreyBgCellStyleHandler;
import com.xingren.excel.handler.impl.RedFontCellStyleHandler;
import com.xingren.excel.handler.impl.YellowBgCellStyleHandler;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/1/18 10:32 上午
 */
@Data
@AllArgsConstructor
@NoArgsConstructor
public class Product {

    public Product(Integer id, Long price, OffsetDateTime created, String name, Boolean isNew, StateEnum stateEnum,
                   LocalDateTime updateTime, String other) {
        this.id = id;
        this.price = price;
        this.created = created;
        this.name = name;
        this.isNew = isNew;
        this.stateEnum = stateEnum;
        this.updateTime = updateTime;
        this.other = other;
    }

    @ExcelColumn(columnName = "id", index = 10)
    private Integer id;

    @ExcelColumn(columnName = "价格", index = 20, centToYuan = true, yuanToCent = true, suffix = " 元",
            columnNameCellStyleHandler =
                    YellowBgCellStyleHandler.class)
    private Long price;

    @ExcelColumn(columnName = "创建日期", index = 60, columnNameCellStyleHandler = GreyBgCellStyleHandler.class)
    private OffsetDateTime created;

    @ExcelColumn(columnName = "名称", index = 40)
    private String name;

    @ExcelColumn(columnName = "是否是新品", index = 41, trueToStr = "新品", falseToStr = "非新品", strToFalse = "非新品",
            strToTrue = "新品")
    private Boolean isNew;

    @ExcelColumn(columnName = "订单状态", index = 50, enumKey = "name", prefix = "状态: ", cellStyleHandler =
            RedFontCellStyleHandler.class)
    private StateEnum stateEnum;

    @ExcelColumn(columnName = "状态变更日期", index = 55)
    private LocalDateTime updateTime;

    @ExcelColumn(columnName = "备注", index = 70, writeConverter = MyConverter.class)
    private String other;

    private String test;

}
