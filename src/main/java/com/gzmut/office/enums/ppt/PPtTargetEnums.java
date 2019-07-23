package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 获取演示文稿中的元素
 * @author livejq
 * @date 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPtTargetEnums {

    /** 幻灯片 */
    SLIDE(1, "幻灯片"),

    /** 文本框 */
    TEXT_BOX(2, "文本框"),

    /** 图片 */
    PICTURE(3, "图片"),

    /** 表格 */
    TABLE(4, "表格"),

    /** SmartArt */
    SMART_ART(5, "SmartArt图形"),

    /** 声音 */
    SOUND(6, "声音"),

    /** 视频 */
    VIDEO(7, "视频");

    private final Integer id;
    private final String target;
}
