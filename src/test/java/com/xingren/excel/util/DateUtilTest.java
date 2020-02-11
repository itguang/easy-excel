package com.xingren.excel.util;

import org.apache.commons.lang3.StringUtils;
import org.junit.Assert;
import org.junit.Test;

import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/2/10 10:56 下午
 */
public class DateUtilTest {

    @Test
    public void testStringOffsetDateTimeConvert() {
        OffsetDateTime offsetDateTime = DateUtil.parseToOffsetDateTime("2020/02/10", "yyyy/MM/dd");
        String time = DateUtil.formatOffsetDateTime(offsetDateTime, "yyyy/MM/dd");
        Assert.assertTrue(time.equals("2020/02/10"));

    }

    @Test
    public void test() {
        String prefix = "";//4
        String value = "10.00";//4
        String postfix = "元";//2

        String text = prefix + value + postfix;//8
        String substring = StringUtils.substring(text, prefix.length(), prefix.length() + value.length());


    }

}
