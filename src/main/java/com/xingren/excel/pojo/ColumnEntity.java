package com.xingren.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;

/**
 * @author guang
 * @since 2020/2/10 6:49 下午
 */
@Getter
@Setter
@AllArgsConstructor
public class ColumnEntity {

    private Cell cell;

    private String columnName;

    private String columnValue;
}
