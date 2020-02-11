package com.xingren.excel.export;

import com.xingren.excel.ExcelConstant;
import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.read.DefaultReadConverter;
import com.xingren.excel.converter.read.EnumReadConverter;
import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.converter.write.*;
import com.xingren.excel.pojo.ColumnEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.pojo.RowEntity;
import com.xingren.excel.util.AnnotationUtil;
import com.xingren.excel.util.DateUtil;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.commons.lang3.StringUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.springframework.core.annotation.AnnotationUtils;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.*;

/**
 * @author guang
 * @since 2020/1/19 4:36 下午
 */
public class ExcelService<T> {

    private ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ExcelService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.forClass(clazz);
    }

    public static <T> ExcelService forClass(Class<T> clazz) {
        return new ExcelService(clazz);
    }

    public Object parseFieldValue(Object rowData, ExcelColumnAnnoEntity entity) {
        String filedName = entity.getFiledName();
        Method getMethod = reflectorUtil.getGetMethod(filedName);
        Class<?> type = entity.getField().getType();

        if (null != entity.getWriteConverter()) {
            Object value = entity.getWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }

        // Enum 类型处理
        if (type.isEnum()) {
            Object value = new EnumWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // OffsetDateTime 类型处理
        if (OffsetDateTime.class.equals(type)) {
            Object value = new OffSetDateTimeWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // LocalDateTime 类型处理
        if (LocalDateTime.class.equals(type)) {
            Object value = new LocalDateTimeWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // Boolean 类型处理
        if (Boolean.class.equals(type)) {
            Object value = new BooleanWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // Integer 和 Long 类型的 分转元处理
        if (entity.getCentToYuan()) {
            Object value = new Cent2YuanWriteConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }

        try {
            Object value = getMethod.invoke(rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return null;

    }

    /**
     * 获取排序好的 注解实体 ExcelColumnAnnoEntity
     */
    public List<ExcelColumnAnnoEntity> getOrderedExcelColumnEntity() {
        List<ExcelColumnAnnoEntity> annoEntities =
                filterExcelField(reflectorUtil.getFieldList()).stream()
                        .map(field -> {
                            ExcelColumnAnnoEntity annoEntity = new ExcelColumnAnnoEntity();
                            annoEntity.setField(field);
                            annoEntity.setFiledName(field.getName());
                            ExcelColumn excelColumn = field.getAnnotation(ExcelColumn.class);
                            Integer index = (Integer) AnnotationUtils.getValue(excelColumn, ANNO_INDEX);
                            String columnName = (String) AnnotationUtils.getValue(excelColumn, ANNO_COLUMN_NAME);
                            String datePattern = (String) AnnotationUtils.getValue(excelColumn, ANNO_DATE_PATTERN);
                            String enumImportKey = (String) AnnotationUtils.getValue(excelColumn, ANNO_ENMUKEY);
                            String trueStr = (String) AnnotationUtils.getValue(excelColumn, ANNO_TRUE_STR);
                            String falseStr = (String) AnnotationUtils.getValue(excelColumn, ANNO_FALSE_STR);
                            String strToFalse = (String) AnnotationUtils.getValue(excelColumn, ANNO_STR_TO_FALSE);
                            String strToTrue = (String) AnnotationUtils.getValue(excelColumn, ANNO_STR_TO_TRUE);
                            Boolean centToYuan = (Boolean) AnnotationUtils.getValue(excelColumn, ANNO_CENT_TO_YUAN);
                            Boolean yuanToCent = (Boolean) AnnotationUtils.getValue(excelColumn, ANNO_YUAN_TO_CENT);
                            String prefix = (String) AnnotationUtils.getValue(excelColumn, ANNO_PREFIX);
                            String suffix = (String) AnnotationUtils.getValue(excelColumn, ANNO_SUFFIX);
                            Class<? extends IWriteConverter> writeConverter =
                                    (Class<? extends IWriteConverter>) AnnotationUtil.getValue(excelColumn,
                                            ANNO_WRITE_CONVERTER);
                            Class<? extends IReadConverter> readConverter =
                                    (Class<? extends IReadConverter>) AnnotationUtil.getValue(excelColumn,
                                            ANNO_READ_CONVERTER);
                            annoEntity.setIndex(index);
                            annoEntity.setColumnName(columnName);
                            annoEntity.setDatePattern(datePattern);
                            annoEntity.setEnumKey(enumImportKey);
                            annoEntity.setTrueStr(trueStr);
                            annoEntity.setFalseStr(falseStr);
                            annoEntity.setStrToFalse(strToFalse);
                            annoEntity.setStrToTrue(strToTrue);
                            annoEntity.setCentToYuan(centToYuan);
                            annoEntity.setYuanToCent(yuanToCent);
                            annoEntity.setSuffix(suffix);
                            annoEntity.setPrefix(prefix);
                            try {
                                if (DefaultWriteConverter.class.equals(writeConverter)) {
                                    annoEntity.setWriteConverter(null);
                                } else {
                                    annoEntity.setWriteConverter(writeConverter.newInstance());
                                }

                                if (DefaultReadConverter.class.equals(readConverter)) {
                                    annoEntity.setReadConverter(null);
                                } else {
                                    annoEntity.setReadConverter(readConverter.newInstance());
                                }

                            } catch (InstantiationException e) {
                                e.printStackTrace();
                            } catch (IllegalAccessException e) {
                                e.printStackTrace();
                            }
                            return annoEntity;
                        }).sorted((o1, o2) -> o1.getIndex().compareTo(o2.getIndex()))
                        .collect(Collectors.toList());
        return annoEntities;
    }

    private List<Field> filterExcelField(List<Field> fieldList) {
        // 过滤出 注解 @ExcelColumn 的字段
        List<Field> excelColumnFields = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class)
                ).collect(Collectors.toList());
        return excelColumnFields;
    }

    public void createSheetHeader(Workbook workbook, int rowIndex, String sheetHeader, Sheet sheet) {
        List<Field> excelField = filterExcelField(reflectorUtil.getFieldList());
        if (CollectionUtils.isEmpty(excelField)) {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, 0));
        } else {
            sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, excelField.size() - 1));
        }
        Row row0 = sheet.createRow(rowIndex);
        Cell header = row0.createCell(0);
        header.setCellValue(sheetHeader);
        header.setCellStyle(ExcelConstant.defaultHeaderStyle(workbook));
        row0.setHeightInPoints(sheetHeaderHeight);
    }

    public <T> List<T> parseRowEntity(ArrayList<RowEntity> rowEntityList) {
        Map<String, Method> fieldNameSetMethodMap = reflectorUtil.getSetMethods();
        Map<String, ExcelColumnAnnoEntity> columnAnnoEntityMap = getColumnAnnoEntityMap();
        List<T> rowDataList = new ArrayList<>(rowEntityList.size());
        for (RowEntity rowEntity : rowEntityList) {
            T rowObj = (T) ReflectorUtil.reflateInstance(clazz);
            for (ColumnEntity columnEntity : rowEntity.getColumnEntityList()) {
                ExcelColumnAnnoEntity annoEntity = columnAnnoEntityMap.get(columnEntity.getColumnName());
                // 如果在 Excel 中的字段在 POJO 中没有相应注解,跳过设值
                if (null == annoEntity) {
                    continue;
                }
                Method setMethod = fieldNameSetMethodMap.get(annoEntity.getFiledName());
                String columnValue = columnEntity.getColumnValue();
                Class<?> fieldType = annoEntity.getField().getType();
                if (null != annoEntity.getReadConverter()) {
                    Object fieldValue = annoEntity.getReadConverter().convert(annoEntity, clazz, columnValue);
                    invokeSetMethod(rowObj, setMethod, fieldValue, fieldType);
                    continue;
                }
                // 前缀/后缀处理
                if (StringUtils.isNotEmpty(annoEntity.getPrefix()) || StringUtils.isNotEmpty(annoEntity.getSuffix())) {
                    columnValue = StringUtils.substring(columnValue, annoEntity.getPrefix().length(),
                            annoEntity.getPrefix().length() + columnValue.length() - annoEntity.getSuffix().length());
                }
                if (String.class.equals(fieldType)) {
                    invokeSetMethod(rowObj, setMethod, columnValue, fieldType);
                    continue;
                }
                // Boolean 类型处理
                if (Boolean.class.equals(fieldType)) {
                    processBoolean(rowObj, annoEntity, setMethod, columnValue, fieldType);
                    continue;
                }
                // 元转分处理
                if (annoEntity.getYuanToCent()) {
                    processYuanToCent(rowObj, setMethod, columnValue, fieldType);
                    continue;
                }
                if (Integer.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
                    invokeSetMethod(rowObj, setMethod, Integer.valueOf(columnValue), fieldType);
                    continue;
                }
                if (Long.class.equals(fieldType) && !annoEntity.getYuanToCent()) {
                    invokeSetMethod(rowObj, setMethod, Long.valueOf(columnValue), fieldType);
                    continue;
                }
                // 枚举处理
                if (fieldType.isEnum()) {
                    // Enum 实例
                    Object enumObj = new EnumReadConverter().convert(annoEntity, clazz, columnValue);
                    invokeSetMethod(rowObj, setMethod, enumObj, fieldType);
                }
                // OffsetDateTime 处理
                if (OffsetDateTime.class.equals(fieldType)) {
                    OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime(columnValue,
                            annoEntity.getDatePattern());
                    invokeSetMethod(rowObj, setMethod, offsetDateTime, fieldType);
                    continue;
                }
                // LocalDateTime 处理
                if (LocalDateTime.class.equals(fieldType)) {
                    LocalDateTime localDateTime = DateUtil.parseToLocalDateTime(columnValue,
                            annoEntity.getDatePattern());
                    invokeSetMethod(rowObj, setMethod, localDateTime, fieldType);
                }
            }
            rowDataList.add(rowObj);
        }
        return rowDataList;
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
        return getOrderedExcelColumnEntity()
                .stream()
                .collect(Collectors.toMap(ExcelColumnAnnoEntity::getColumnName, Function.identity()));
    }

}
