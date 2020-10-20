/**
 * Copyright (c) 2018, biezhi (biezhi.me@gmail.com)
 * <p>
 * Licensed under the Apache License, Version 2.0 (the "License");
 * you may not use this file except in compliance with the License.
 * You may obtain a copy of the License at
 * <p>
 * http://www.apache.org/licenses/LICENSE-2.0
 * <p>
 * Unless required by applicable law or agreed to in writing, software
 * distributed under the License is distributed on an "AS IS" BASIS,
 * WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
 * See the License for the specific language governing permissions and
 * limitations under the License.
 */
package com.xingren.excel;

import org.apache.poi.ss.usermodel.*;

/**
 * @author guang
 * @since 2020/1/17 2:42 下午
 */
public interface ExcelConstant {

    String ANNO_COLUMN_NAME = "columnName";
    String ANNO_INDEX = "index";
    String ANNO_DATE_PATTERN = "datePattern";
    String ANNO_ENMUKEY = "enumKey";
    String ANNO_WRITE_CONVERTER = "writeConverter";
    String ANNO_READ_CONVERTER = "readConverter";
    String ANNO_CELL_STYLE_HANDLER = "cellStyleHandler";
    String ANNO_COLUMN_CELL_STYLE_HANDLER = "columnNameCellStyleHandler";
    String ANNO_TRUE_STR = "trueToStr";
    String ANNO_FALSE_STR = "falseToStr";
    String ANNO_STR_TO_TRUE = "strToTrue";
    String ANNO_STR_TO_FALSE = "strToFalse";
    String ANNO_CENT_TO_YUAN = "centToYuan";
    String ANNO_YUAN_TO_CENT = "yuanToCent";
    String ANNO_REQUIRED = "required";
    String ANNO_SUFFIX = "suffix";
    String ANNO_PREFIX = "prefix";

    /**
     * The default worksheet name.
     */
    String DEFAULT_SHEET_NAME = "Sheet0";
    String DEFAULT_FONT_NAME = "SimHei";
    int DEFAULT_COLUMN_WIDTH = 20 * 256;

    String DEFAULT_DATE_PATTREN = "yyyy/MM/dd";

    String ERROR_COLUMN_NAME = "错误信息";

    /**
     * sheet header 列高
     */
    int sheetHeaderRowHeight = 30;

    /**
     * sheet header 字体大小
     */
    int sheetHeaderRowFontSize = 15;

    /**
     * columnName row 高
     */
    int columnTitleRowHeight = 25;
    /**
     *
     */
    int columnDataRowHeight = 20;

    /**
     * 默认的 Excel Header 样式
     */
    static CellStyle defaultHeaderRowStyle(Workbook workbook) {
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setWrapText(true);
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);
        font.setFontHeightInPoints((short) sheetHeaderRowFontSize);
        font.setBold(true);
        headerCellStyle.setFont(font);
        return headerCellStyle;
    }

    /**
     * 默认的 Excel Tail 样式
     */
    static CellStyle defaultTailRowStyle(Workbook workbook) {
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setAlignment(HorizontalAlignment.LEFT);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setWrapText(true);
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);
        font.setBold(false);
        headerCellStyle.setFont(font);
        return headerCellStyle;
    }

    /**
     * 默认的 Excel ColumnName 样式
     */
    static CellStyle defaultColumnNameStyle(Workbook workbook) {
        CellStyle headerCellStyle = workbook.createCellStyle();
        headerCellStyle.setAlignment(HorizontalAlignment.CENTER);
        headerCellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerCellStyle.setWrapText(true);
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);
//        font.setBold(true);
        headerCellStyle.setFont(font);
        return headerCellStyle;
    }

    /**
     * 默认的 Exlel Row 数据样式,垂直水平居中
     */
    static CellStyle defaultDataRowStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);
        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);
        cellStyle.setFont(font);
        return cellStyle;
    }

}
