package com.xingren.excel.util;

import org.apache.commons.lang3.StringUtils;

/**
 * @author guang
 * @since 2020/2/27 10:39 上午
 */
public class StrUtil {

    /**
     * 将字符串的首字母转大写
     *
     * @param str 需要转换的字符串
     * @return
     */
    public static String upperFirstChar(String str) {
        if (StringUtils.isEmpty(str)) {
            throw new RuntimeException("字符串不能为空");
        }
        // 进行字母的ascii编码前移，效率要高于截取字符串进行转换的操作
        char[] cs = str.toCharArray();
        if (Character.isUpperCase(cs[0])) {
            return str;
        }
        cs[0] -= 32;
        return String.valueOf(cs);
    }
}
