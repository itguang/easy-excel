package com.xingren.excel.pojo;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.usermodel.Cell;

/**
 * 单个Cell 实体类
 *
 * @author guang
 * @since 2020/2/10 6:49 下午
 */
@Getter
@Setter
@AllArgsConstructor
public class CellEntity {

    /**
     * 当前 Cell 对象
     */
    private Cell cell;

    /**
     * 注解字段名称
     */
    private String cellName;

    /**
     * 导入 excel 中的 cell 值
     */
    private String cellValue;
}
