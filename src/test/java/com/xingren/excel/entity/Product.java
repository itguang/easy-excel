package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.handler.impl.GreyBgCellStyleHandler;
import com.xingren.excel.handler.impl.RedFontCellStyleHandler;
import com.xingren.excel.handler.impl.YellowBgCellStyleHandler;
import lombok.AllArgsConstructor;
import lombok.Data;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/1/18 10:32 上午
 */
@Data
@AllArgsConstructor
public class Product {

    @ExcelColumn(columnName = "id", index = 10)
    private Integer id;

    @ExcelColumn(columnName = "价格", index = 20, centToYuan = true, suffix = " 元", columnNameCellStyleHandler =
            YellowBgCellStyleHandler.class)
    private Long price;

    @ExcelColumn(columnName = "创建日期", index = 60, columnNameCellStyleHandler = GreyBgCellStyleHandler.class)
    private OffsetDateTime created;

    @ExcelColumn(columnName = "名称", index = 40)
    private String name;

    @ExcelColumn(columnName = "是否是新品", index = 41, trueToStr = "新品", falseToStr = "非新品")
    private Boolean isNew;

    @ExcelColumn(columnName = "订单状态", index = 50, enumKey = "name", prefix = "状态: ", cellStyleHandler =
            RedFontCellStyleHandler.class)
    private StateEnum stateEnum;

    @ExcelColumn(columnName = "状态变更日期", index = 55)
    private LocalDateTime updateTime;

    @ExcelColumn(columnName = "备注", index = 70)
    private String other;

    // 导入必须有无参构造器
    public Product() {
    }
}
