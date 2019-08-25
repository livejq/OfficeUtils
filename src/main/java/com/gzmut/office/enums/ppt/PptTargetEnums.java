package com.gzmut.office.enums.ppt;

import com.gzmut.office.enums.CommonEnum;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 获取演示文稿中的目标元素
 *
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PptTargetEnums implements CommonEnum {

    /**
     * 节
     */
    SECTION(1, "SECTION", "节"),

    /**
     * 幻灯片
     */
    SLIDE(2, "SLIDE", "幻灯片"),

    /**
     * 文本框
     */
    TEXT_BOX(3, "TEXT_BOX", "文本框"),

    /**
     * 图片
     */
    PICTURE(4, "PICTURE", "图片"),

    /**
     * 表格
     */
    TABLE(5, "TABLE", "表格"),

    /**
     * SmartArt
     */
    SMART_ART(6, "SMART_ART", "SmartArt"),

    /**
     * 声音
     */
    SOUND(7, "SOUND", "声音"),

    /**
     * 视频
     */
    VIDEO(8, "VIDEO", "视频"),

    /**
     * 图表
     */
    CHART(9, "CHART", "图表"),

    /**
     * 动画
     */
    ANIMATION(10, "ANIMATION", "动画"),

    /**
     * 形状
     */
    SHAPE(10, "SHAPE", "形状");

    /**
     * 字段ID
     */
    private final Integer code;

    /**
     * 目标字段
     */
    private final String Name;

    /**
     * 字段描述
     */
    private final String desc;
}