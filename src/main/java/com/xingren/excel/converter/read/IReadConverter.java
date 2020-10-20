package com.xingren.excel.converter.read;

import com.xingren.excel.converter.Converter;
import com.xingren.excel.pojo.CellEntity;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * Excel 读取转换器
 * <note>converter 都会放在循环中,容易出现性能瓶颈,注意缓存</note>
 *
 * @param <T> rowObject 对应的 java 类型
 * @param <D> 要转成的目标类型
 */
public interface IReadConverter<T, D> extends Converter {

    /**
     * @param annoEntity Column 注解 实体
     * @param cellEntity cellEntity <note>CellEntity#cellValue 没有前缀和后缀</note>
     * @param rowObj     rowObj
     * @return 转换后的值
     */
    D convert(ExcelColumnAnnoEntity annoEntity, CellEntity cellEntity, T rowObj);
}
