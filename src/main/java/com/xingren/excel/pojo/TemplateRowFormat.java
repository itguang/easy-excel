package com.xingren.excel.pojo;

import com.xingren.excel.enums.FormatType;
import lombok.Getter;
import lombok.Setter;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * @author ellin
 * @since 2020/08/03
 */
@Getter
@Setter
public class TemplateRowFormat {

    private CellRangeAddress range;
    private FormatType formatType;
    private String[] data;

    public TemplateRowFormat(int firstRow, int firstCol, int lastRow, int lastCol) {
        this.range = new CellRangeAddress(firstRow, lastRow, firstCol, lastCol);
    }

}
