package com.xingren.excel.util;

import com.xingren.excel.annotation.ExcelColumn;
import com.xingren.excel.entity.Product;
import com.xingren.excel.entity.ReflectorEntity;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;

public class ReflectorUtilTest {
    private static ReflectorUtil productReflectorUtil;

    @BeforeAll
    public static void setUp() {
        productReflectorUtil = ReflectorUtil.fromCache(Product.class);

    }

    @Test
    public void test_reflect_entity() {
        ReflectorUtil myReflectorUtil = ReflectorUtil.forClass(ReflectorEntity.class);
    }

    @Test
    public void testFilterExcelColumnField() {

        List<Field> fieldList = productReflectorUtil.getFieldList();

        List<Field> excelColumnFields = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class)
                ).collect(Collectors.toList());

        assertEquals(excelColumnFields.size(), fieldList.size() - 1);

    }

    @Test
    public void test_getMethods() {

        Map<String, Method> getMethods = productReflectorUtil.getGetMethods();

        assertNotNull(getMethods.get("id"));
        assertNotNull(getMethods.get("price"));
        assertNotNull(getMethods.get("created"));
        assertNotNull(getMethods.get("name"));
    }

    @Test
    public void test_setMethods() {
        Map<String, Method> setMethods = productReflectorUtil.getSetMethods();

        assertNotNull(setMethods.get("id"));
        assertNotNull(setMethods.get("price"));
        assertNotNull(setMethods.get("created"));
        assertNotNull(setMethods.get("name"));

    }

}