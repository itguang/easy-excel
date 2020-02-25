package com.xingren.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author guang
 * @since 2020/2/10 6:49 下午
 */
@Data
@AllArgsConstructor
public class ColumnEntity {

    private Cell cell;

    private String columnName;

    private String columnValue;
}
