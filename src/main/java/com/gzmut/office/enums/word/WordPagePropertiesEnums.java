package com.gzmut.office.enums.word;

/**
 * 页面属性枚举类
 * @author MXDC
 * @date 2019/7/16
 */
public enum WordPagePropertiesEnums {

    /** 页面大小*/
   PAGE_SIZE(1,"pageSize"),
    /** 页面方向*/
    PAGA_ORI(2,"page");



    /* 属性id**/
    private Integer id;
    /* 属性名称**/
    private String property;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getProperty() {
        return property;
    }

    public void setProperty(String property) {
        this.property = property;
    }

    WordPagePropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
