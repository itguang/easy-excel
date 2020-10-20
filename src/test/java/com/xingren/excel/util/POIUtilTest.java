package com.xingren.excel.util;

import org.junit.jupiter.api.Assertions;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

public class POIUtilTest {

    @Test
    @DisplayName("计算 Cell 自动高度")
    void getExcelCellAutoHeight() {
        // 中文
        Assertions.assertEquals(Float.valueOf(24.0f), POIUtil.getExcelCellAutoHeight("测试", 2f));
        // 英文
        Assertions.assertEquals(Float.valueOf(12.0f), POIUtil.getExcelCellAutoHeight("test", 4f));
        // 全角逗号
        Assertions.assertEquals(Float.valueOf(24.0f), POIUtil.getExcelCellAutoHeight("，", 1f));
        // 全角符号和中文
        Assertions.assertEquals(Float.valueOf(24.0f), POIUtil.getExcelCellAutoHeight("测试，", 3f));
    }
}