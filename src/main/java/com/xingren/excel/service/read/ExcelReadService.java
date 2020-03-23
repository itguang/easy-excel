package com.xingren.excel.service.read;

import com.xingren.excel.converter.read.impl.EnumReadConverter;
import com.xingren.excel.exception.ExcelException;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ErrorInfoRow;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.pojo.RowEntity;
import com.xingren.excel.service.ExcelColumnService;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.lang3.ObjectUtils;
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

        // 如果在 Excel 中的字段在 POJO 中没有相应注解,跳过赋值
        if (null == annoEntity) {
            return;
        }
        Method setMethod = fieldNameSetMethodMap.get(annoEntity.getFiledName());
        String columnValue = cellEntity.getCellValue();
        Class<?> fieldType = annoEntity.getField().getType();
        if (null != annoEntity.getReadConverter()) {
            Object fieldValue = annoEntity.getReadConverter().convert(annoEntity, clazz, columnValue, rowObj);
            invokeSetMethod(rowObj, setMethod, fieldValue, fieldType, annoEntity, columnValue);
            return;
        }
        // 前缀/后缀处理
        if (StringUtils.isNotEmpty(annoEntity.getPrefix()) || StringUtils.isNotEmpty(annoEntity.getSuffix())) {
            columnValue = StringUtils.substring(columnValue, annoEntity.getPrefix().length(),
                    annoEntity.getPrefix().length() + columnValue.length() - annoEntity.getSuffix().length());
        }
        if (String.class.equals(fieldType)) {
            invokeSetMethod(rowObj, setMethod, columnValue, fieldType, annoEntity, columnValue);
            return;
        }
        // Boolean 类型处理
        if (Boolean.class.equals(fieldType)) {
            processBoolean(rowObj, annoEntity, setMethod, columnValue, fieldType);
            return;
        }
        // 元转分处理
        if (annoEntity.getYuanToCent()) {
            processYuanToCent(rowObj, setMethod, columnValue, fieldType, annoEntity);
            return;
        }
        if (Integer.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
            invokeSetMethod(rowObj, setMethod, Integer.valueOf(columnValue), fieldType, annoEntity, columnValue);
            return;
        }
        if (Long.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
            invokeSetMethod(rowObj, setMethod, Long.valueOf(columnValue), fieldType, annoEntity, columnValue);
            return;
        }
        // 枚举处理
        if (fieldType.isEnum()) {
            // Enum 实例
            Object enumObj = new EnumReadConverter().convert(annoEntity, clazz, columnValue, rowObj);
            invokeSetMethod(rowObj, setMethod, enumObj, fieldType, annoEntity, columnValue);
            return;
        }
        // OffsetDateTime 处理
        if (OffsetDateTime.class.equals(fieldType)) {
            if (StringUtils.isNotEmpty(columnValue)) {
                Cell cell = cellEntity.getCell();
                OffsetDateTime offsetDateTime = null;
                if (cell.getCellTypeEnum().equals(NUMERIC)) {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(columnValue, DEFAULT_DATE_PATTREN);
                } else {
                    offsetDateTime = DateUtil.parseToOffsetDateTime(columnValue, annoEntity.getDatePattern());
                }
                invokeSetMethod(rowObj, setMethod, offsetDateTime, fieldType, annoEntity, columnValue);
            }
            return;
        }
        // LocalDateTime 处理
        if (LocalDateTime.class.equals(fieldType)) {
            Cell cell = cellEntity.getCell();
            LocalDateTime localDateTime = null;
            if (cell.getCellTypeEnum().equals(NUMERIC)) {
                localDateTime = DateUtil.parseToLocalDateTime(columnValue, DEFAULT_DATE_PATTREN);
            } else {
                localDateTime = DateUtil.parseToLocalDateTime(columnValue, annoEntity.getDatePattern());
            }
            invokeSetMethod(rowObj, setMethod, localDateTime, fieldType, annoEntity, columnValue);
            return;
        }
    }

    private <T> void processBoolean(T rowObj, ExcelColumnAnnoEntity annoEntity, Method setMethod, String columnValue,
                                    Class<?> fieldType) {
        String strToFalse = annoEntity.getStrToFalse();
        String strToTrue = annoEntity.getStrToTrue();
        if (StringUtils.isNotEmpty(strToFalse) && columnValue.equals(strToFalse) || "false".equals(columnValue)) {
            invokeSetMethod(rowObj, setMethod, Boolean.FALSE, fieldType, annoEntity, columnValue);
            return;
        }
        if (StringUtils.isNotEmpty(strToTrue) && columnValue.equals(strToTrue) || "true".equals(columnValue)) {
            invokeSetMethod(rowObj, setMethod, Boolean.TRUE, fieldType, annoEntity, columnValue);
            return;
        }
    }

    private <T> void processYuanToCent(T rowObj, Method setMethod, String columnValue, Class<?> fieldType,
                                       ExcelColumnAnnoEntity annoEntity) {
        if (StringUtils.isBlank(columnValue)) {
            return;
        }
        Double doubleValue = Double.valueOf(columnValue) * 100;
        if (Integer.class.equals(fieldType)) {
            Integer intValue = doubleValue.intValue();
            invokeSetMethod(rowObj, setMethod, intValue, fieldType, annoEntity, columnValue);

        }
        if (Long.class.equals(fieldType)) {
            Long longValue = doubleValue.longValue();
            invokeSetMethod(rowObj, setMethod, longValue, fieldType, annoEntity, columnValue);
        }
    }

    private <T> void invokeSetMethod(T rowObj, Method setMethod, Object setValue, Class<?> setClass,
                                     ExcelColumnAnnoEntity annoEntity, String columnValue) {
        try {
            Class<?> fieldType = annoEntity.getField().getType();
            // required 为 true 处理
            if (annoEntity.getRequired()) {
                // 枚举值不合法
                if (fieldType.isEnum() && StringUtils.isNotBlank(columnValue)) {
                    throw new ExcelException("枚举值[" + columnValue + "]不合法");
                }
                // 空值
                if (ObjectUtils.isEmpty(setValue) || "".equals(setValue)) {
                    throw new ExcelException("[" + annoEntity.getColumnName() + "]不能为空");
                }

            } else {
                // required 为 false 处理
                // 如果是枚举类型
                if (fieldType.isEnum()) {
                    if (null == setValue && StringUtils.isNotBlank(columnValue)) {
                        throw new ExcelException("枚举值[" + columnValue + "]不合法");
                    }
                    if (null == setValue) {
                        return;
                    }
                }
            }
            ReflectorUtil.invokeSetMethod(rowObj, setMethod, setValue, setClass);
        } catch (Exception e) {
            // 是否需要出错信息进行捕获
            if (isErrorInfoRow(rowObj)) {
                setErrotInfo((ErrorInfoRow) rowObj, e.getMessage());
                e.printStackTrace();
            } else {
                throw new ExcelException(e.getMessage());
            }
        }
    }

    private <T> void setErrotInfo(ErrorInfoRow rowObj, String errorInfo) {
        rowObj.setErrorInfo(errorInfo + "\n");
    }

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
