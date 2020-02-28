package com.xingren.excel.util;

/**
 * @author guang
 * @since 2020/2/8 3:14 下午
 */

import org.apache.commons.lang3.StringUtils;

import java.text.DateFormat;
import java.text.ParseException;
import java.text.SimpleDateFormat;
import java.time.Instant;
import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.time.ZoneId;
import java.time.format.DateTimeFormatter;
import java.util.Date;

/**
 * 日期时间工具类
 */
public class DateUtil {

    /**
     * 格式化时间
     *
     * @param dateTime
     * @param patten
     * @return
     */
    public static String formatOffsetDateTime(OffsetDateTime dateTime, String patten) {
        if (dateTime == null || StringUtils.isEmpty(patten)) {
            return null;
        }

        return dateTime.format(DateTimeFormatter.ofPattern(patten));
    }

    public static String formatDate(Date time, String patten) {
        DateFormat dateFormat = new SimpleDateFormat(patten);
        return dateFormat.format(time);
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

    /**
     * 根据字符串设置日期
     *
     * @param text
     * @param datePattern
     * @return
     */
    public static OffsetDateTime parseToOffsetDateTime(String text, String datePattern) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(datePattern);
        Date date = null;
        try {
            date = simpleDateFormat.parse(text);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        Instant instant = date.toInstant();
        ZoneId zone = ZoneId.systemDefault();

        return OffsetDateTime.ofInstant(instant, zone);
    }

    public static LocalDateTime parseToLocalDateTime(String text, String datePattern) {
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat(datePattern);
        Date date = null;
        try {
            date = simpleDateFormat.parse(text);
            System.out.println(date);
        } catch (ParseException e) {
            e.printStackTrace();
        }

        Instant instant = date.toInstant();
        ZoneId zone = ZoneId.systemDefault();
        return LocalDateTime.ofInstant(instant, zone);

    }

}
