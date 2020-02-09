package com.xingren.excel.export;

import com.xingren.excel.ExcelConstant;
import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.*;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.collections.CollectionUtils;
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
import java.util.List;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.*;

/**
 * @author guang
 * @since 2020/1/19 4:36 下午
 */
public class ExcelExportService {

    private ReflectorUtil reflectorUtil;

    private Class clazz;

    private ExcelExportService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.forClass(clazz);
    }

    public static ExcelExportService forClass(Class clazz) {
        return new ExcelExportService(clazz);
    }

    public Object parseFieldValue(Object rowData, ExcelColumnAnnoEntity entity) {
        String filedName = entity.getFiledName();
        Method getMethod = reflectorUtil.getGetMethod(filedName);
        Class<?> type = entity.getField().getType();

        if (null != entity.getConverter()) {
            Object value = entity.getConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }

        // Enum 类型处理
        if (type.isEnum()) {
            Object value = new EnumConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // OffsetDateTime 类型处理
        if (OffsetDateTime.class.equals(type)) {
            Object value = new OffSetDateTimeConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // LocalDateTime 类型处理
        if (LocalDateTime.class.equals(type)) {
            Object value = new LocalDateTimeConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // Boolean 类型处理
        if (Boolean.class.equals(type)) {
            Object value = new BooleanConverter().convert(entity, clazz, rowData);
            return entity.getPrefix() + value + entity.getSuffix();
        }
        // Integer 和 Long 类型的 分转元处理
        if (entity.getCentToYuan()) {
            Object value = new Cent2YuanConverter().convert(entity, clazz, rowData);
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
                            Boolean centToYuan = (Boolean) AnnotationUtils.getValue(excelColumn, ANNO_CENT_TO_YUAN);
                            String prefix = (String) AnnotationUtils.getValue(excelColumn, ANNO_PREFIX);
                            String suffix = (String) AnnotationUtils.getValue(excelColumn, ANNO_SUFFIX);
                            Class<? extends IConverter> converter =
                                    (Class<? extends IConverter>) AnnotationUtils.getValue(excelColumn, ANNO_CONVERTER);
                            annoEntity.setIndex(index);
                            annoEntity.setColumnName(columnName);
                            annoEntity.setDatePattern(datePattern);
                            annoEntity.setEnumKey(enumImportKey);
                            annoEntity.setTrueStr(trueStr);
                            annoEntity.setFalseStr(falseStr);
                            annoEntity.setCentToYuan(centToYuan);
                            annoEntity.setSuffix(suffix);
                            annoEntity.setPrefix(prefix);
                            try {
                                if (DefaultConverter.class.equals(converter)) {
                                    annoEntity.setConverter(null);
                                } else {
                                    annoEntity.setConverter(converter.newInstance());
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

}
