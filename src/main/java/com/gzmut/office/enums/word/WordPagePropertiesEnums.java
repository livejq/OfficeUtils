package com.gzmut.office.enums.word;

/**
 * 页面属性枚举类
 * @author MXDC
 * @date 2019/7/16
 */
public enum WordPagePropertiesEnums {

    /** 页面大小*/
    PAGE_SIZE(1,"页面大小"),
    /** 页面方向*/
    PAGA_ORIENT(2,"页面方向"),
    /** 上页边距*/
    PAGE_MARGIN_TOP(3,"上页边距"),
    /** 下页边距*/
    PAGE_MARGIN_BOTTOM(4,"下页边距"),
    /** 左页边距*/
    PAGE_MARGIN_LEFT(5,"左页边距"),
    /** 右页边距*/
    PAGE_MARGIN_RIGHT(6,"右页边距"),
    /** 页眉距边界*/
    PAGE_MARGIN_HEADER(7,"页眉距边界"),
    /** 页脚距边界*/
    PAGE_MARGIN_FOOTER(8,"页脚距边界");



    /* 属性id*/
    private Integer id;
    /* 属性名称*/
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
