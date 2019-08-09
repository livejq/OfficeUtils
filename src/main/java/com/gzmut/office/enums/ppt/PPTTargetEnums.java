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
public enum PPTTargetEnums {

    /** 幻灯片 */
    SLIDE(1, "SLIDE"),

    /** 文本框 */
    TEXT_BOX(2, "TEXT_BOX"),

    /** 图片 */
    PICTURE(3, "PICTURE"),

    /** 表格 */
    TABLE(4, "TABLE"),

    /** SmartArt */
    SMART_ART(5, "SMART_ART"),

    /** 声音 */
    SOUND(6, "SOUND"),

    /** 视频 */
    VIDEO(7, "VIDEO"),

    /** 图表 */
    CHART(8, "CHART"),

    /** 动画 */
    ANIMATION(9, "ANIMATION");

    private final Integer id;
    private final String target;
}
