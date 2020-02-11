package com.xingren.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Data;

import java.util.List;

/**
 * @author guang
 * @since 2020/2/10 7:01 下午
 */
@Data
@AllArgsConstructor
public class RowEntity {

    private List<ColumnEntity> columnEntityList;

}
