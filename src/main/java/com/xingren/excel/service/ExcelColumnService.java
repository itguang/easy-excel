package com.xingren.excel.service;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.converter.read.IReadConverter;
import com.xingren.excel.converter.read.impl.DefaultReadConverter;
import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.converter.write.impl.DefaultWriteConverter;
import com.xingren.excel.handler.ICellStyleHandler;
import com.xingren.excel.handler.IColumnNameCellStyleHandler;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.AnnotationUtil;
import com.xingren.excel.util.ReflectorUtil;
import org.apache.poi.ss.formula.functions.T;
import org.springframework.core.annotation.AnnotationUtils;

import java.lang.reflect.Field;
import java.util.List;
import java.util.stream.Collectors;

import static com.xingren.excel.ExcelConstant.*;

/**
 * @author guang
 * @since 2020/2/11 4:14 下午
 */
public class ExcelColumnService {

    private ReflectorUtil reflectorUtil;

    private Class<T> clazz;

    private ExcelColumnService(Class clazz) {
        this.clazz = clazz;
        this.reflectorUtil = ReflectorUtil.forClass(clazz);
    }

    public static <T> ExcelColumnService forClass(Class<T> clazz) {
        return new ExcelColumnService(clazz);
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
                            Class<? extends ICellStyleHandler> cellStyleHandler =
                                    (Class<? extends ICellStyleHandler>) AnnotationUtil.getValue(excelColumn,
                                            ANNO_CELL_STYLE_HANDLER);
                            Class<? extends IColumnNameCellStyleHandler> columnCellStyleHandler =
                                    (Class<? extends IColumnNameCellStyleHandler>) AnnotationUtil.getValue(excelColumn,
                                            ANNO_COLUMN_CELL_STYLE_HANDLER);
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
                                annoEntity.setCellStyleHandler(cellStyleHandler.newInstance());
                                annoEntity.setColumnNameCellStyleHandler(columnCellStyleHandler.newInstance());
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

}
