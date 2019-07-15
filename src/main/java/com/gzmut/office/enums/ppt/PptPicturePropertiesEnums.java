package com.gzmut.office.enums.ppt;

/**
 * 封装图片属性
 * @author livejq
 * @date 2019/7/13
 **/
public enum PptPicturePropertiesEnums {

    /*
     *
     * 根据id获取图片属性
     * 尺寸
     **/
    SIZE(1, "尺寸"),

    /*
     *
     * 根据id获取图片属性
     * 超链接
     **/
    HYPERLINK(2, "超链接"),

    /*
     *
     * 根据id获取图片属性
     * 类型
     **/
    CONTENT_TYPE(3, "类型"),

    /*
     *
     * 根据id获取图片属性
     * 后缀
     **/
    EXTENSION(4, "后缀");


    private Integer id;
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

    PptPicturePropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
