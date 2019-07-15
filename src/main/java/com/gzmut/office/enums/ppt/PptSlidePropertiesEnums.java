package com.gzmut.office.enums.ppt;

/**
 * 封装幻灯片属性
 * @author livejq
 * @date 2019/7/13
 **/
public enum PptSlidePropertiesEnums {

    /*
     *
     * 根据id获取幻灯片属性
     * 背景颜色
     **/
    BACKGROUND_COLOR(1, "背景颜色"),

    /*
     *
     * 根据id获取幻灯片属性
     * 切入方式
     **/
    SLIP_STYLE(2, "切入方式"),

    /*
     *
     * 根据id获取幻灯片属性
     * 主题
     **/
    THEME(3, "主题"),

    /*
     *
     * 根据id获取幻灯片属性
     * 布局类型
     **/
    LAYOUT_STYLE(4, "布局类型"),

    /*
     *
     * 根据id获取幻灯片属性
     * 尺寸
     **/
    SIZE(5, "尺寸"),

    /*
     *
     * 根据id获取幻灯片属性
     * 标题
     **/
    TITLE(6, "标题");


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

    PptSlidePropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
