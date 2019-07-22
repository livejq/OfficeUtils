package com.gzmut.office.enums.word;

/**
 * Word表格属性
 */
public enum WordTablePropertiesEnums {

    //表格宽
    TABLE_WIDTH(1,"表格宽度"),
    //表格高
    TABLE_HEIGHT(2,"表格高度");



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

    WordTablePropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
