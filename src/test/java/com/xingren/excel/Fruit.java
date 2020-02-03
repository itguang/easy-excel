package com.xingren.excel;

import java.time.OffsetDateTime;

/**
 * @author guang
 * @since 2020/1/18 10:32 上午
 */
public class Fruit {

    private Integer id;

    private Long price;

    private OffsetDateTime created;

    private String name;

    private StateEnum stateEnum;

    private String other;

    public Fruit(Integer id, Long price, OffsetDateTime created, String name) {
        this.id = id;
        this.price = price;
        this.created = created;
        this.name = name;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public Long getPrice() {
        return price;
    }

    public void setPrice(Long price) {
        this.price = price;
    }

    public OffsetDateTime getCreated() {
        return created;
    }

    public void setCreated(OffsetDateTime created) {
        this.created = created;
    }

    public String getName() {
        return name;
    }

    public void setName(String name) {
        this.name = name;
    }

    public StateEnum getStateEnum() {
        return stateEnum;
    }

    public void setStateEnum(StateEnum stateEnum) {
        this.stateEnum = stateEnum;
    }

    public String other() {
        return this.other;
    }
}
