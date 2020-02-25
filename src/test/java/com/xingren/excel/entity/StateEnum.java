package com.xingren.excel.entity;

public enum StateEnum {

    UP(0, "上架"),
    DOWN(1, "下架"),
    NONE(2, "库存不足");

    StateEnum(int code, String name) {
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
