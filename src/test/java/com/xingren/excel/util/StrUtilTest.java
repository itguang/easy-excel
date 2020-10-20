package com.xingren.excel.util;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.Test;

public class StrUtilTest {

    @Test
    public void test_captureName() {
        Assertions.assertTrue(StrUtil.upperFirstChar("name").equals("Name"));
        Assertions.assertTrue(StrUtil.upperFirstChar("Name").equals("Name"));
    }
}