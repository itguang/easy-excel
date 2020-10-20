package com.xingren.excel.util;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

import java.time.LocalDateTime;
import java.time.OffsetDateTime;
import java.util.Date;

/**
 * @author guang
 * @since 2020/2/10 10:56 下午
 */
public class DateUtilTest {

    @Test
    public void test_stringOffsetDateTimeConvert() {
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime("2020/02/10", "yyyy/MM/dd");
        String time = DateUtil.formatOffsetDateTime(offsetDateTime, "yyyy/MM/dd");
        Assertions.assertTrue(time.equals("2020/02/10"));

    }

    @Test
    public void test_parseToOffsetDateTime() {
        LocalDateTime localDateTime = DateUtil.parseToLocalDateTime("2020/02/10", "yyyy/MM/dd");
        String time = DateUtil.formatLocalDateTime(localDateTime, "yyyy/MM/dd");
        Assertions.assertEquals("2020/02/10", time);

    }

    @Test
    public void test_dateToLocalDateTime() {
        Date date = new Date();
        LocalDateTime localDateTime = DateUtil.dateToLocalDateTime(date);
        Date parseDate = DateUtil.localDateTimeToDate(localDateTime);
        Assertions.assertEquals(date.getTime(), parseDate.getTime());

    }

}
