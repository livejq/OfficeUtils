package com.gzmut.office.enums.word;

/**
 * 封装判题方法
 * @author MXDC
 * @date 2019/7/16
 **/
public enum WordCorrectEnums {

    /** 判断文件是否存在*/
    CHECK_FILE_IS_EXIST(1,"checkFileIsExist");


    /* 属性id**/
    private Integer id;
    /* 检查方法名称**/
    private String checkName;

    private WordCorrectEnums(Integer id, String checkName) {
        this.id = id;
        this.checkName = checkName;
    }

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getCheckName() {
        return checkName;
    }

    public void setCheckName(String checkName) {
        this.checkName = checkName;
    }


}
