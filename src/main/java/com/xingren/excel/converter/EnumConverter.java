package com.xingren.excel.converter;

import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * @author guang
 * @since 2020/2/8 11:35 上午
 */
public class EnumConverter implements IConverter {

    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Class<?> clazz, Object rowData) {
        Object value = null;

        String filedName = entity.getFiledName();
        Method getMethod = ReflectorUtil.forClass(clazz).getGetMethod(filedName);

        Object enumObj = null;
        try {
            enumObj = getMethod.invoke(rowData);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }

        Class<?> enumClass = entity.getField().getType();
        if (!enumClass.isEnum()) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:" + filedName + " 不是枚举类型!");
        }

        Method[] declaredMethods = enumClass.getDeclaredMethods();
        String enumKeyMethodName = "get" + captureName(entity.getEnumKey());
        Optional<Method> enumKeyMethod = Stream.of(declaredMethods).filter(method -> {
            return method.getName().equals(enumKeyMethodName);
        }).findFirst();

        if (!enumKeyMethod.isPresent()) {
            throw new ExcelConvertException("方法 " + enumKeyMethodName
                    + " 在 " + enumClass.getName() + " 不存在");
        }

        try {
            value = enumKeyMethod.get().invoke(enumObj);
        } catch (IllegalAccessException e) {
            e.printStackTrace();
        } catch (InvocationTargetException e) {
            e.printStackTrace();
        }
        return value;
    }

    /**
     * 将字符串的首字母转大写
     *
     * @param str 需要转换的字符串
     * @return
     */
    private static String captureName(String str) {
        // 进行字母的ascii编码前移，效率要高于截取字符串进行转换的操作
        char[] cs = str.toCharArray();
        cs[0] -= 32;
        return String.valueOf(cs);
    }
}
