package com.gzmut.office.enums.word;

/**
 * Word表格单元格属性
 */
public enum WordCellPropertiesEnums {

    //单元格宽
    CELL_WIDTH(1,"单元格宽度"),
    //单元格高
    CELL_HEIGHT(2,"单元格高度");



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

    WordCellPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }
}
