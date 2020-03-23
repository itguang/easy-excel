package com.xingren.excel.pojo;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.handler.ICellStyleHandler;
import lombok.Getter;
import lombok.Setter;

import java.lang.reflect.Field;

/**
 * @author guang
 * @since 2020/2/7 12:18 下午
 */
@Getter
@Setter
public class ExcelColumnAnnoEntity {

    private Field field;

    private String filedName;

    private String columnName;

    private Integer index;

    private String datePattern;

    private String enumKey;

    private IWriteConverter writeConverter;

    private IReadConverter readConverter;

    private ICellStyleHandler cellStyleHandler;

    private ICellStyleHandler columnNameCellStyleHandler;

    private String trueStr;

    private String falseStr;

    private String strToFalse;

    private String strToTrue;

    private Boolean centToYuan;

    private Boolean yuanToCent;

    private Boolean required;

    private String prefix;

    private String suffix;

}
