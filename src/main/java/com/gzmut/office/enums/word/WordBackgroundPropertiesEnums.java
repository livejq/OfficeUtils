package com.gzmut.office.enums.word;

/**
 * 背景属性
 * @author MXDC
 * @date 2019/7/16
 */
public enum  WordBackgroundPropertiesEnums {

    /** 背景颜色*/
    BACKGROUND_COLOR(1,"background_color"),
    /** 背景填充效果--渐变*/
    BACKGROUND_FILL_GRADIENT(2,"background_fill_gradient"),
    /** 背景填充效果--纹理*/
    BACKGROUND_FILL_TILE(3,"background_fill_tile"),
    /** 背景填充效果--图案*/
    BACKGROUND_FILL_PATTERN(4,"background_fill_pattern");

    /** 背景填充效果--图片*/
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

     WordBackgroundPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
