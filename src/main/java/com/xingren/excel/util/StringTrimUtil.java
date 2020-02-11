package com.xingren.excel.util;

/**
 * @author guang
 * @since 2020/2/11 11:11 上午
 */
public class StringTrimUtil {
    public static String custom_ltrim(String str, char c) {
        char[] chars = str.toCharArray();
        int len = chars.length;
        int st = 0;
        while ((st < len) && (chars[st] == c)) {
            st++;
        }
        return (st > 0) ? str.substring(st, len) : str;
    }

    public static String custom_rtrim(String str, char c) {
        char[] chars = str.toCharArray();
        int len = chars.length;
        int st = 0;
        while ((st < len) && (chars[len - 1] == c)) {
            len--;
        }
        return (len < chars.length) ? str.substring(st, len) : str;
    }
}
