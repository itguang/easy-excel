package com.xingren.excel.util;

import java.util.regex.Pattern;

/**
 * @author guang
 * @since 2020/9/4 2:19 下午
 */
public class POIUtil {

    private static Pattern wordPattern = Pattern.compile("^[A-Za-z0-9]+$");
    private static Pattern fullWidthPattern = Pattern.compile("[\u4e00-\u9fa5]+$");
    private static Pattern cnPattern = Pattern.compile("[^x00-xff]");

    /**
     * 根据 cell 的值,计算 cell 高度
     *
     * @param cellValue       cell 值
     * @param fontCountInline 该单元格每行多少个汉字 全角为1 英文或符号为0.5
     * @return 合适的 cell 宽度值
     */
    public static Float getExcelCellAutoHeight(String cellValue, Float fontCountInline) {
        //每一行的高度指定
        float defaultRowHeight = 12.00f;
        float defaultCount = 0.00f;
        for (int i = 0; i < cellValue.length(); i++) {
            float ff = getRegex(cellValue.substring(i, i + 1));
            defaultCount = defaultCount + ff;
        }
        return ((int) (defaultCount / fontCountInline) + 1) * defaultRowHeight;
    }

    private static Float getRegex(String charStr) {
        if (charStr == " ") {
            return 0.5f;
        }
        // 判断是否为字母或字符
        if (wordPattern.matcher(charStr).matches()) {
            return 0.5f;
        }
        // 判断是否为全角
        if (fullWidthPattern.matcher(charStr).matches()) {
            return 1.00f;
        }
        //全角符号 及中文

        if (cnPattern.matcher(charStr).matches()) {
            return 1.00f;
        }
        return 0.5f;
    }
}