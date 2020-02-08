package com.xingren.excel.pojo;

import com.xingren.excel.converter.IConverter;
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

    private IConverter converter;

    private String trueStr;

    private String falseStr;

    private Boolean centToYuan;

    private String prefix;

    private String suffix;

    public Field getField() {
        return field;
    }

    public void setField(Field field) {
        this.field = field;
    }

    public String getFiledName() {
        return filedName;
    }

    public void setFiledName(String filedName) {
        this.filedName = filedName;
    }

    public String getColumnName() {
        return columnName;
    }

    public void setColumnName(String columnName) {
        this.columnName = columnName;
    }

    public Integer getIndex() {
        return index;
    }

    public void setIndex(Integer index) {
        this.index = index;
    }

    public String getDatePattern() {
        return datePattern;
    }

    public void setDatePattern(String datePattern) {
        this.datePattern = datePattern;
    }

    public String getEnumKey() {
        return enumKey;
    }

    public void setEnumKey(String enumKey) {
        this.enumKey = enumKey;
    }

    public IConverter getConverter() {
        return converter;
    }

    public void setConverter(IConverter converter) {
        this.converter = converter;
    }

    public String getTrueStr() {
        return trueStr;
    }

    public void setTrueStr(String trueStr) {
        this.trueStr = trueStr;
    }

    public String getFalseStr() {
        return falseStr;
    }

    public void setFalseStr(String falseStr) {
        this.falseStr = falseStr;
    }

    public Boolean getCentToYuan() {
        return centToYuan;
    }

    public void setCentToYuan(Boolean centToYuan) {
        this.centToYuan = centToYuan;
    }
}
