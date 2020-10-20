package com.xingren.excel.converter;

import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

/**
 * @author guang
 * @since 2020/8/20 2:23 下午
 */
public interface ConverterFactory {

    /**
     * 根据注解值获取对应的 Converter
     *
     * @param annoEntity {@link ExcelColumnAnnoEntity}
     * @return Converter
     */
    Converter getConverter(ExcelColumnAnnoEntity annoEntity);

}
