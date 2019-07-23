package com.gzmut.office.enums.word;

/**
 * 封装word字体属性
 * @author MXDC
 * @date 2019/7/16
 **/
public enum  WordFontPropertiesEnums {

    /**
     * 字体大小
     */
    SIZE(1,"大小"),

    /** 字体颜色*/
    COLOR(2,"颜色");



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

     WordFontPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
