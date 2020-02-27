package com.xingren.excel.util;

import org.junit.Assert;
import org.junit.Test;

public class StrUtilTest {

    @Test
    public void test_captureName() {
        Assert.assertTrue(StrUtil.upperFirstChar("name").equals("Name"));
        Assert.assertTrue(StrUtil.upperFirstChar("Name").equals("Name"));
    }
}