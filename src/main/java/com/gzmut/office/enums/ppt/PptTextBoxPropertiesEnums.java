package com.gzmut.office.enums.ppt;

/**
 * 封装文本框属性
 * @author livejq
 * @date 2019/7/13
 **/
public enum PptTextBoxPropertiesEnums {

    /*
     *
     * 根据id获取文本框属性
     * 填充颜色
     **/
    FILL_COLOR(1, "填充颜色"),

    /*
     *
     * 根据id获取文本框属性
     * 文本内容
     **/
    TEXT_CONTENT(2, "文本内容"),

    /*
     *
     * 根据id获取文本框属性
     * 字体样式
     **/
    TEXT_STYLE(3, "字体样式"),

    /*
     *
     * 根据id获取文本框属性
     * 字体大小
     **/
    TEXT_SIZE(4, "字体大小"),

    /*
     *
     * 根据id获取文本框属性
     * 字体颜色
     **/
    TEXT_COLOR(5, "字体颜色"),

    /*
     *
     * 根据id获取文本框属性
     * 超链接
     **/
    HYPERLINK(6, "超链接"),

    /*
     *
     * 根据id获取文本框属性
     * 粗体
     **/
    BOLD(7, "粗体"),

    /*
     *
     * 根据id获取文本框属性
     * 斜体
     **/
    ITALIC(8, "斜体"),

    /*
     *
     * 根据id获取文本框属性
     * 下划线
     **/
    UNDERLINED(9, "下划线"),

    /*
     *
     * 根据id获取文本框属性
     * 脚注
     **/
    SUBSCRIPT(10, "脚注"),

    /*
     *
     * 根据id获取文本框属性
     * 段前间距
     **/
    SPACE_BEFORE(11, "段前间距"),

    /*
     *
     * 根据id获取文本框属性
     * 段后间距
     **/
    SPACE_AFTER(12, "段后间距"),

    /*
     *
     * 根据id获取文本框属性
     * 行高
     **/
    LINE_SPACE(13, "行高");


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

    PptTextBoxPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
