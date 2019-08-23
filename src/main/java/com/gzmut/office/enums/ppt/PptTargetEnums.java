package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 获取演示文稿中的目标元素
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PptTargetEnums {

    /** 节 */
    SECTION(1, "SECTION"),

    /** 幻灯片 */
    SLIDE(2, "SLIDE"),

    /** 文本框 */
    TEXT_BOX(3, "TEXT_BOX"),

    /** 图片 */
    PICTURE(4, "PICTURE"),

    /** 表格 */
    TABLE(5, "TABLE"),

    /** SmartArt */
    SMART_ART(6, "SMART_ART"),

    /** 声音 */
    SOUND(7, "SOUND"),

    /** 视频 */
    VIDEO(8, "VIDEO"),

    /** 图表 */
    CHART(9, "CHART"),

    /** 动画 */
    ANIMATION(10, "ANIMATION");

    private final Integer id;
    private final String target;
}
