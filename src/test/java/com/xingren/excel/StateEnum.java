package com.xingren.excel;

public enum StateEnum {

    /**
     * 应收款批次未处理
     */
    NONE(0, "未处理"),
    /**
     * 应收款批次处理中
     */
    DOING(1, "处理中"),
    /**
     * 应收款批次处理已经完成
     */
    COMPLETED(2, "已完成");

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

}
