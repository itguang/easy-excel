package com.xingren.excel.util;

/**
 * @author guang
 * @since 2020/2/8 3:14 下午
 */

import org.apache.commons.lang3.StringUtils;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.ZoneOffset;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * 日期时间工具类
 */
public class DateUtil {

    private static final String OFFSET_ID = "+08:00";

    /**
     * 根据字符串设置日期
     *
     * @param text
     * @param formatter
     * @return
     */
    private static OffsetDateTime parse(CharSequence text, DateTimeFormatter formatter) {
        if (StringUtils.isEmpty(text) || formatter == null) {
            return null;
        }
        LocalDateTime localDateTime = LocalDateTime.parse(text, formatter);
        OffsetDateTime dateTime = OffsetDateTime.of(localDateTime, ZoneOffset.of(OFFSET_ID));
        return dateTime;
    }

    /**
     * 格式化时间
     *
     * @param dateTime
     * @param patten
     * @return
     */
    public static String formatDateTime(OffsetDateTime dateTime, String patten) {
        if (dateTime == null || StringUtils.isEmpty(patten)) {
            return null;
        }

        return dateTime.format(DateTimeFormatter.ofPattern(patten));
    }

    /**
     * Date转LocalDateTime
     *
     * @param date Date对象
     * @return
     */
    public static LocalDateTime dateToLocalDateTime(Date date) {
        return LocalDateTime.ofInstant(date.toInstant(), ZoneId.systemDefault());
    }

    /**
     * LocalDateTime转换为Date
     *
     * @param dateTime LocalDateTime对象
     * @return
     */
    public static Date localDateTimeToDate(LocalDateTime dateTime) {
        return Date.from(dateTime.atZone(ZoneId.systemDefault()).toInstant());
    }

    /**
     * 按pattern格式化时间-默认yyyy-MM-dd HH:mm:ss格式
     *
     * @param dateTime LocalDateTime对象
     * @param pattern  要格式化的字符串
     * @return
     */
    public static String formatLocalDateTime(LocalDateTime dateTime, String pattern) {
        if (dateTime == null || StringUtils.isEmpty(pattern)) {
            return null;
        }

        DateTimeFormatter formatter = DateTimeFormatter.ofPattern(pattern);
        return dateTime.format(formatter);
    }

}
