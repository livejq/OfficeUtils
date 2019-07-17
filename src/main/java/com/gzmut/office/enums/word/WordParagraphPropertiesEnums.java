package com.gzmut.office.enums.word;

/**
 * 段落属性
 * @author MXDC
 * @date 2019/7/17
 **/
public enum  WordParagraphPropertiesEnums {

    /** 对齐方式*/
    ALIGNMENT(1,"alignment"),
    /** 段落缩进方式*/
    INDENTATION(2,"indentation");


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

    WordParagraphPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }

}
