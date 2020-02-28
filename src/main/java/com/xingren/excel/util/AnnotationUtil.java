package com.xingren.excel.util;

import java.lang.annotation.Annotation;
import java.lang.reflect.Method;

/**
 * @author guang
 * @since 2020/2/11 11:40 上午
 */
public class AnnotationUtil {

    public static Object getValue(Annotation annotation, String attributeName) {
        try {
            Method method = annotation.annotationType().getDeclaredMethod(attributeName);
            ReflectorUtil.makeAccessible(method);
            return method.invoke(annotation);
        } catch (Exception ex) {
            throw new RuntimeException("注解 " + annotation.getClass().getName() + "找不到属性: " + attributeName);
        }
    }
}
