package com.xingren.excel.pojo;

import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.converter.write.IWriteConverter;
import lombok.Data;

import java.lang.reflect.Field;

/**
 * @author guang
 * @since 2020/2/7 12:18 下午
 */
@Data
public class ExcelColumnAnnoEntity {

    private Field field;

    private String filedName;

    private String columnName;

    private Integer index;

    private String datePattern;

    private String enumKey;

    private IWriteConverter writeConverter;

    private IReadConverter readConverter;

    private String trueStr;

    private String falseStr;

    private String strToFalse;

    private String strToTrue;

    private Boolean centToYuan;

    private Boolean yuanToCent;

    private String prefix;

    private String suffix;

}
