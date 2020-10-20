package com.xingren.excel.converter.write.impl;

import com.xingren.excel.converter.write.IWriteConverter;
import com.xingren.excel.exception.ExcelConvertException;
import com.xingren.excel.pojo.ExcelColumnAnnoEntity;
import com.xingren.excel.util.ReflectorUtil;
import com.xingren.excel.util.StrUtil;
import org.apache.commons.lang3.StringUtils;

import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.util.Optional;
import java.util.stream.Stream;

/**
 * @author guang
 * @since 2020/2/8 11:35 上午
 */
public class EnumWriteConverter implements IWriteConverter {

    @Override
    public Object convert(ExcelColumnAnnoEntity entity, Object rowData) {
        Class<?> clazz = rowData.getClass();
        String filedName = entity.getFiledName();
        Method getMethod = ReflectorUtil.fromCache(clazz).getGetMethod(filedName);

        Class<?> enumClass = entity.getField().getType();
        if (!enumClass.isEnum()) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中字段:" + filedName + " 不是枚举类型!");
        }
        if (StringUtils.isEmpty(entity.getEnumKey())) {
            throw new ExcelConvertException("类 " + clazz.getName() + " 中枚举字段:" + filedName + " 没有指定 enumKey 值!");
        }

        Object enumObj = null;
        try {
            enumObj = getMethod.invoke(rowData);
        } catch (IllegalAccessException | InvocationTargetException e) {
            throw new RuntimeException(e);
        }

        Method[] declaredMethods = enumClass.getDeclaredMethods();
        String enumKeyMethodName = "get" + StrUtil.upperFirstChar(entity.getEnumKey());
        Optional<Method> enumKeyMethod = Stream.of(declaredMethods).filter(method -> {
            return method.getName().equals(enumKeyMethodName);
        }).findFirst();

        if (!enumKeyMethod.isPresent()) {
            throw new ExcelConvertException("方法 " + enumKeyMethodName
                    + " 在 " + enumClass.getName() + " 不存在");
        }

        return ReflectorUtil.invokeGetMethod(enumKeyMethod.get(), enumObj);
    }

}
