package com.xingren.excel.util;

import com.xingren.excel.Fruit;
import org.junit.Before;
import org.junit.Test;

import java.lang.reflect.Method;
import java.util.Map;

import static org.junit.Assert.assertNotNull;
import static org.junit.Assert.assertNull;

public class ReflectorUtilTest {
    ReflectorUtil reflectorUtil;

    @Before
    public void setUp() throws Exception {
        reflectorUtil = ReflectorUtil.forClass(Fruit.class);
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