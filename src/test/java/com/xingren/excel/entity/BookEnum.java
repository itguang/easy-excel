package com.xingren.excel.entity;

/**
 * @author guang
 * @since 2020/3/18 3:08 下午
 */
public enum BookEnum {

    UP(0, "上架"),
    DOWN(1, "下架");

    BookEnum(int code, String name) {
        this.code = code;
        this.name = name;
    }

    /**
     * code
     */
    private int code;
    /**
     * name
     */
    private String name;

    public int getCode() {
        return code;
    }

    public String getName() {
        return name;
    }

}
