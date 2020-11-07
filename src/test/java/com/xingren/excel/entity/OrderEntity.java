package com.xingren.excel.entity;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.pojo.ErrorInfoRow;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.EqualsAndHashCode;
import lombok.NoArgsConstructor;

/**
 * @author guang
 * @since 2020/3/14 10:44 上午
 */
@AllArgsConstructor
@NoArgsConstructor
@EqualsAndHashCode(callSuper=false)
@Data
public class OrderEntity extends ErrorInfoRow {

    @ExcelColumn(index = 10, columnName = "订单ID")
    private Integer id;

    @ExcelColumn(index = 20, columnName = "订单号")
    private String orderNo;
}
