package com.xingren.excel.service.write;

import com.xingren.excel.ExcelConstant;
import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.write.impl.*;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.commons.collections.CollectionUtils;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;

import java.lang.reflect.Field;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.List;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.sheetHeaderRowHeight;

/**
 * @author guang
 * @since 2020/1/19 4:36 下午
 */
public class ExcelWriteService<T> {

    private ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ExcelWriteService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.forClass(clazz);
    }

    public static <T> ExcelWriteService forClass(Class<T> clazz) {
        return new ExcelWriteService(clazz);
    }

    public Object parseFieldValue(Object rowData, ExcelColumnAnnoEntity entity, Cell cell) {
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
        header.setCellStyle(ExcelConstant.defaultHeaderRowStyle(workbook));
        row0.setHeightInPoints(sheetHeaderRowHeight);
    }

}
