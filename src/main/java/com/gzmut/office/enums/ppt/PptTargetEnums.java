package com.gzmut.office.enums.ppt;

/**
 * 封装目标对象
 * @author livejq
 * @date 2019/7/13
 **/
public enum PptTargetEnums {
    /*
     *
     * 根据id获取目标对象名称
     * 获取幻灯片
     **/
    SLIDE(1, "slide"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取文本框
     **/
    TEXT_BOX(2, "text_box"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取图片
     **/
    PICTURE(3, "picture"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取表格
     **/
    TABLE(4, "table"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取SmartArt
     **/
    SMART_ART(5, "smart_art"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取声音
     **/
    SOUND(6, "sound"),

    /*
     *
     * 根据id获取目标对象名称
     * 获取视频
     **/
    VIDEO(7, "video");

    private Integer id;
    private String target;

    public Integer getId() {
        return id;
    }

    public void setId(Integer id) {
        this.id = id;
    }

    public String getTarget() {
        return target;
    }

    public void setTarget(String target) {
        this.target = target;
    }

    PptTargetEnums(Integer id, String target) {
        this.id = id;
        this.target = target;
    }
}
