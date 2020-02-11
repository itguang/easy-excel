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
            return null;
        }
    }
}
