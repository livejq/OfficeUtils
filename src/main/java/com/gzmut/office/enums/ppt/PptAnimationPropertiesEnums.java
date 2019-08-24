package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 根据id获取图片属性
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PptAnimationPropertiesEnums {

    /** 动画类型 */
    ANIMATION_TYPE(1, "ANIMATION_TYPE"),

    /** 触发动画方式 */
    ANIMATION_START_STYLE(2, "ANIMATION_START_STYLE"),

    /** 动画延迟时间（ms） */
    DELAY(3, "DELAY"),

    /** 动画持续时间（ms) */
    DUR(4, "DUR");

    private final Integer id;
    private final String property;
}
