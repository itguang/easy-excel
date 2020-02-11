package com.xingren.excel.converter.read;

import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.HashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.Collectors;
import java.util.stream.Stream;

/**
 * @author guang
 * @since 2020/2/11 2:44 下午
 */
public class EnumReadConverter implements IReadConverter {
    private static Map<Class, Map<String, Object>> classEnmuObj = new HashMap<>();

    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object cellValue) {
        Class<?> fieldType = entity.getField().getType();
        if (!fieldType.isEnum()) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:"
                    + entity.getFiledName() + " 不是 Enum 类型!");
        }
        return getEnumConstant(fieldType).get(cellValue);
    }

    public Map<String, Object> getEnumConstant(Class<?> type) {
        if (!classEnmuObj.isEmpty()) {
            return classEnmuObj.get(type);
        }
        List<Method> getMethods = Stream.of(type.getDeclaredMethods())
                .filter(method -> method.getName().startsWith("get"))
                .collect(Collectors.toList());

        Map<String, Object> map = new HashMap<>(getMethods.size());
        if (type.isEnum()) {
            Object[] enumConstants = type.getEnumConstants();
            for (Object obj : enumConstants) {
                for (Method method : getMethods) {
                    try {
                        Object value = type.getMethod(method.getName()).invoke(obj);
                        map.put(value.toString(), obj);
                    } catch (IllegalAccessException e) {
                        e.printStackTrace();
                    } catch (InvocationTargetException e) {
                        e.printStackTrace();
                    } catch (NoSuchMethodException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
        classEnmuObj.put(type, map);
        return map;
    }
}
