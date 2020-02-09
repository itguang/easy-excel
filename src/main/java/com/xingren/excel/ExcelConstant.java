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
    String ANNO_CONVERTER = "converter";
    String ANNO_TRUE_STR = "trueToStr";
    String ANNO_FALSE_STR = "falseToStr";
    String ANNO_CENT_TO_YUAN = "centToYuan";
    String ANNO_SUFFIX = "suffix";
    String ANNO_PREFIX = "prefix";

    /**
     * The default worksheet name.
     */
    String DEFAULT_SHEET_NAME = "Sheet0";
    String DEFAULT_FONT_NAME = "SimHei";
    int DEFAULT_COLUMN_WIDTH = 20 * 256;

    /**
     * 列头高度
     */
    int sheetHeaderHeight = 30;

    /**
     *
     */
    int columnTitleHeight = 20;
    static CellStyle defaultRowTitleStyle(Workbook workbook) {
        CellStyle style = workbook.createCellStyle();
        style.setAlignment(HorizontalAlignment.CENTER);
        style.setVerticalAlignment(VerticalAlignment.CENTER);
        return style;
    }

    /**
     * The default Excel header style.
     *
     * @param workbook Excel workbook
     * @return header row cell style
     */
    static CellStyle defaultHeaderStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);

        cellStyle.setDataFormat((short) 0);

        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);

        cellStyle.setFont(font);
        return cellStyle;
    }

    /**
     * The default Excel column style.
     *
     * @param workbook Excel workbook
     * @return row column cell style
     */
    static CellStyle defaultColumnStyle(Workbook workbook) {
        CellStyle cellStyle = workbook.createCellStyle();
        cellStyle.setAlignment(HorizontalAlignment.CENTER);
        cellStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        cellStyle.setWrapText(true);

        cellStyle.setDataFormat((short) 0);

        Font font = workbook.createFont();
        font.setFontName(DEFAULT_FONT_NAME);

        cellStyle.setFont(font);
        return cellStyle;
    }

}
