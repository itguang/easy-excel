package com.xingren.excel.service.read;

import com.xingren.excel.converter.read.ReadConverterFactory;
import com.xingren.excel.converter.read.impl.PrefixAndSuffixReadConverter;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ErrorInfoRow;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.pojo.RowEntity;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;

import java.lang.reflect.Method;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

/**
 * @author guang
 * @since 2020/2/11 3:30 下午
 */
public class ExcelReadService {

    private ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ReadConverterFactory readConverterFactory = new ReadConverterFactory();

    private ExcelReadService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.fromCache(clazz);
    }

    public static <T> ExcelReadService forClass(Class<T> clazz) {
        return new ExcelReadService(clazz);
    }

    public <T> List<T> parseRowEntity(ArrayList<RowEntity> rowEntityList) {
        Map<String, ExcelColumnAnnoEntity> columnNameAnnoEntityMap = getColumnAnnoEntityMap();
        List<T> rowDataList = new ArrayList<>(rowEntityList.size());
        for (RowEntity rowEntity : rowEntityList) {
            T rowObj = (T) ReflectorUtil.reflateInstance(clazz);
            for (CellEntity cellEntity : rowEntity.getCellEntityList()) {
                ExcelColumnAnnoEntity annoEntity = columnNameAnnoEntityMap.get(cellEntity.getCellName());
                parseColumnEntityToRowObj(annoEntity, rowObj, cellEntity);
            }
            rowDataList.add(rowObj);
        }
        return rowDataList;
    }

    private <T> void parseColumnEntityToRowObj(ExcelColumnAnnoEntity annoEntity, T rowObj, CellEntity cellEntity) {
        Map<String, Method> fieldNameSetMethodMap = reflectorUtil.getSetMethods();

        if (null == annoEntity) {
            return;
        }

        Method setMethod = fieldNameSetMethodMap.get(annoEntity.getFiledName());
        Class<?> fieldType = annoEntity.getField().getType();
        String cellValue = null;
        // 前缀/后缀处理(去除前缀后缀)
        if (StringUtils.isNotEmpty(annoEntity.getPrefix()) || StringUtils.isNotEmpty(annoEntity.getSuffix())) {
            cellValue = new PrefixAndSuffixReadConverter().convert(annoEntity, cellEntity, rowObj);
            // 去除前缀后缀后更新 cellValue 值
            cellEntity.setCellValue(cellValue);
        }

        // 以下基本数据类型的处理 ↓↓↓↓↓

        Object value = readConverterFactory.getConverter(annoEntity).convert(annoEntity, cellEntity, rowObj);
        invokeSetMethod(rowObj, setMethod, value, fieldType, annoEntity, cellValue);
    }

    /**
     * @param rowObj     当前 row 对象的 java 对象
     * @param setMethod  当前 字段的 setXXX 方法
     * @param setValue   当前字段值(已经有具体类型的值)
     * @param setClass   当前字段类型
     * @param annoEntity 当前字段上的注解
     * @param cellValue  当前字段值(统一为 String 类型的形式)
     * @return 范湖赋值后的 row 对象
     */
    private <T> void invokeSetMethod(T rowObj, Method setMethod, Object setValue, Class<?> setClass,
                                     ExcelColumnAnnoEntity annoEntity, String cellValue) {
        try {
            Class<?> fieldType = annoEntity.getField().getType();
            // required 为 true 处理: 分为 枚举值 和 非枚举值 两种情况
            if (annoEntity.getRequired()) {
                // 枚举值不合法
                if (fieldType.isEnum() && null == setValue) {
                    throw new ExcelException("[" + annoEntity.getColumnName() + "]不合法");
                }
                // 空值
                if (null == setValue || "".equals(setValue)) {
                    throw new ExcelException("[" + annoEntity.getColumnName() + "]不能为空");
                }

            } else {
                // required 为 false 处理
                // 如果是枚举类型
                if (fieldType.isEnum()) {
                    if (null == setValue && StringUtils.isNotBlank(cellValue)) {
                        throw new ExcelException("[" + cellValue + "]不合法");
                    }
                    if (null == setValue) {
                        return;
                    }
                }
            }
            ReflectorUtil.invokeSetMethod(rowObj, setMethod, setValue, setClass);
        } catch (Exception e) {
            // 是否需要对出错信息进行捕获
            if (isErrorInfoRow(rowObj)) {
                setErrorInfo((ErrorInfoRow) rowObj, e.getMessage());
                // 把异常信息栈打印出来,但是不主动抛出,因为继承了 ErrorInfoRow 类
                e.printStackTrace();
            } else {
                throw new ExcelException(e.getMessage());
            }
        }
    }

    /**
     * 设置出错信息,多个用换行符分割
     */
    private <T> void setErrorInfo(ErrorInfoRow rowObj, String errorInfo) {
        rowObj.setErrorInfo(errorInfo + "\n");
    }

    /**
     * 判断 该对象是否继承自 ErrorInfoRow
     */
    private <T> boolean isErrorInfoRow(T rowObj) {
        return rowObj instanceof ErrorInfoRow;
    }

    private Map<String, ExcelColumnAnnoEntity> getColumnAnnoEntityMap() {
        return ExcelColumnService.forClass(clazz)
                .getOrderedExcelColumnEntity()
                .stream()
                .collect(Collectors.toMap(ExcelColumnAnnoEntity::getColumnName, Function.identity()));
    }

}
