package com.gzmut.office.enums.word;

/**
 * @author MXDC
 * @date 2019/7/22
 **/
public enum WordBrodesPropertiesEnums {
    ee(1,"2");


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

    WordBrodesPropertiesEnums(Integer id, String property) {
        this.id = id;
        this.property = property;
    }

}
