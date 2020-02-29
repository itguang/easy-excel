package com.xingren.excel.service.read;

import com.xingren.excel.converter.read.impl.EnumReadConverter;
import com.xingren.excel.pojo.ColumnEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.pojo.RowEntity;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.formula.functions.T;
import org.apache.poi.ss.usermodel.Cell;

import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.DEFAULT_DATE_PATTREN;
import static org.apache.poi.ss.usermodel.CellType.NUMERIC;

/**
 * @author guang
 * @since 2020/2/11 3:30 下午
 */
public class ExcelReadService {

    private ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ExcelReadService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.fromCache(clazz);
    }

    public static <T> ExcelReadService forClass(Class<T> clazz) {
        return new ExcelReadService(clazz);
    }

    public <T> List<T> parseRowEntity(ArrayList<RowEntity> rowEntityList) {
        Map<String, ExcelColumnAnnoEntity> columnAnnoEntityMap = getColumnAnnoEntityMap();
        List<T> rowDataList = new ArrayList<>(rowEntityList.size());
        for (RowEntity rowEntity : rowEntityList) {
            T rowObj = (T) ReflectorUtil.reflateInstance(clazz);
            for (ColumnEntity columnEntity : rowEntity.getColumnEntityList()) {
                if (!columnEntity.getCell().toString().equals("")) {
                    parseColumnEntityToRowObj(columnAnnoEntityMap, rowObj, columnEntity);
                }
            }
            rowDataList.add(rowObj);
        }
        return rowDataList;
    }

    private <T> void parseColumnEntityToRowObj(Map<String, ExcelColumnAnnoEntity> columnAnnoEntityMap,
                                               T rowObj, ColumnEntity columnEntity) {
        Map<String, Method> fieldNameSetMethodMap = reflectorUtil.getSetMethods();
        ExcelColumnAnnoEntity annoEntity = columnAnnoEntityMap.get(columnEntity.getColumnName());
        // 如果在 Excel 中的字段在 POJO 中没有相应注解,跳过设值
        if (null == annoEntity) {
            return;
        }
        Method setMethod = fieldNameSetMethodMap.get(annoEntity.getFiledName());
        String columnValue = columnEntity.getColumnValue();
        Class<?> fieldType = annoEntity.getField().getType();
        if (null != annoEntity.getReadConverter()) {
            Object fieldValue = annoEntity.getReadConverter().convert(annoEntity, clazz, columnValue);
            invokeSetMethod(rowObj, setMethod, fieldValue, fieldType);
            return;
        }
        // 前缀/后缀处理
        if (StringUtils.isNotEmpty(annoEntity.getPrefix()) || StringUtils.isNotEmpty(annoEntity.getSuffix())) {
            columnValue = StringUtils.substring(columnValue, annoEntity.getPrefix().length(),
                    annoEntity.getPrefix().length() + columnValue.length() - annoEntity.getSuffix().length());
        }
        if (String.class.equals(fieldType)) {
            invokeSetMethod(rowObj, setMethod, columnValue, fieldType);
            return;
        }
        // Boolean 类型处理
        if (Boolean.class.equals(fieldType)) {
            processBoolean(rowObj, annoEntity, setMethod, columnValue, fieldType);
            return;
        }
        // 元转分处理
        if (annoEntity.getYuanToCent()) {
            processYuanToCent(rowObj, setMethod, columnValue, fieldType);
            return;
        }
        if (Integer.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
            invokeSetMethod(rowObj, setMethod, Integer.valueOf(columnValue), fieldType);
            return;
        }
        if (Long.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
            invokeSetMethod(rowObj, setMethod, Long.valueOf(columnValue), fieldType);
            return;
        }
        // 枚举处理
        if (fieldType.isEnum()) {
            // Enum 实例
            Object enumObj = new EnumReadConverter().convert(annoEntity, clazz, columnValue);
            invokeSetMethod(rowObj, setMethod, enumObj, fieldType);
        }
        // OffsetDateTime 处理
        if (OffsetDateTime.class.equals(fieldType)) {
            if (StringUtils.isNotEmpty(columnValue)) {
                Cell cell = columnEntity.getCell();
                OffsetDateTime offsetDateTime = null;
                if (cell.getCellTypeEnum().equals(NUMERIC)) {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(columnValue, DEFAULT_DATE_PATTREN);
                } else {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(columnValue, annoEntity.getDatePattern());
                }
                invokeSetMethod(rowObj, setMethod, offsetDateTime, fieldType);
            }
            return;
        }
        // LocalDateTime 处理
        if (LocalDateTime.class.equals(fieldType)) {
            Cell cell = columnEntity.getCell();
            LocalDateTime localDateTime = null;
            if (cell.getCellTypeEnum().equals(NUMERIC)) {
                localDateTime = DateUtil.parseToLocalDateTime(columnValue, DEFAULT_DATE_PATTREN);
            } else {
                localDateTime = DateUtil.parseToLocalDateTime(columnValue, annoEntity.getDatePattern());
            }
            invokeSetMethod(rowObj, setMethod, localDateTime, fieldType);
        }
    }

    private <T> void processBoolean(T rowObj, ExcelColumnAnnoEntity annoEntity, Method setMethod, String columnValue,
                                    Class<?> fieldType) {
        String strToFalse = annoEntity.getStrToFalse();
        String strToTrue = annoEntity.getStrToTrue();
        if (StringUtils.isNotEmpty(strToFalse) && columnValue.equals(strToFalse) || "false".equals(columnValue)) {
            invokeSetMethod(rowObj, setMethod, Boolean.FALSE, fieldType);
            return;
        }
        if (StringUtils.isNotEmpty(strToTrue) && columnValue.equals(strToTrue) || "true".equals(columnValue)) {
            invokeSetMethod(rowObj, setMethod, Boolean.TRUE, fieldType);
            return;
        }
    }

    private <T> void processYuanToCent(T rowObj, Method setMethod, String columnValue, Class<?> fieldType) {
        Double doubleValue = Double.valueOf(columnValue) * 100;
        if (Integer.class.equals(fieldType)) {
            Integer intValue = doubleValue.intValue();
            invokeSetMethod(rowObj, setMethod, intValue, fieldType);

        }
        if (Long.class.equals(fieldType)) {
            Long longValue = doubleValue.longValue();
            invokeSetMethod(rowObj, setMethod, longValue, fieldType);

        }
    }

    private <T> void invokeSetMethod(T rowObj, Method setMethod, Object setValue, Class<?> setClass) {
        ReflectorUtil.invokeSetMethod(rowObj, setMethod, setValue, setClass);
    }

    private Map<String, ExcelColumnAnnoEntity> getColumnAnnoEntityMap() {
        return ExcelColumnService.forClass(clazz)
                .getOrderedExcelColumnEntity()
                .stream()
                .collect(Collectors.toMap(ExcelColumnAnnoEntity::getColumnName, Function.identity()));
    }

}
