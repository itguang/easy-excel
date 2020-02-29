package com.xingren.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.List;

/**
 * @author guang
 * @since 2020/2/10 7:01 下午
 */
@Getter
@Setter
@AllArgsConstructor
public class RowEntity {

    private List<ColumnEntity> columnEntityList;

}
