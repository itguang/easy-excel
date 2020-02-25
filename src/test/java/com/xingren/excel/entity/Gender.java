package com.xingren.excel.entity;

public enum Gender {

    MALE(0, "男"),
    FEMALE(1, "女");

    Gender(int code, String name) {
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

    public void setCode(int code) {
        this.code = code;
    }

    public java.lang.String getName() {
        return name;
    }

    public void setName(java.lang.String name) {
        this.name = name;
    }

}
