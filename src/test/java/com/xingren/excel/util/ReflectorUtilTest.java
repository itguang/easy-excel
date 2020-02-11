package com.xingren.excel.util;

import com.xingren.excel.Product;
import com.xingren.excel.annotation.ExcelColumn;
import org.junit.Before;
import org.junit.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Method;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;

import static org.junit.Assert.*;

public class ReflectorUtilTest {
    ReflectorUtil reflectorUtil;

    @Before
    public void setUp() throws Exception {
        reflectorUtil = ReflectorUtil.forClass(Product.class);
    }

    @Test
    public void testFilterExcelColumnField() {

        List<Field> fieldList = reflectorUtil.getFieldList();

        List<Field> excelColumnFields = fieldList.stream()
                .filter(field -> field.isAnnotationPresent(ExcelColumn.class)
                ).collect(Collectors.toList());

        assertEquals(excelColumnFields.size(), fieldList.size() - 1);

    }

    @Test
    public void test_getMethods() {

        Map<String, Method> getMethods = reflectorUtil.getGetMethods();

        assertNotNull(getMethods.get("id"));
        assertNotNull(getMethods.get("price"));
        assertNotNull(getMethods.get("created"));
        assertNotNull(getMethods.get("name"));
        assertNull(getMethods.get("other"));
    }

    @Test
    public void test_setMethods() {
        Map<String, Method> setMethods = reflectorUtil.getSetMethods();

        assertNotNull(setMethods.get("id"));
        assertNotNull(setMethods.get("price"));
        assertNotNull(setMethods.get("created"));
        assertNotNull(setMethods.get("name"));
        assertNull(setMethods.get("other"));

    }



}