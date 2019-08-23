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

    /** 进入动画 */
    IN(1, "IN"),

    /** 强调动画 */
    EMPHASIZE(2, "EMPHASIZE"),

    /** 退出动画 */
    OUT(3, "OUT"),

    /** 单击触发动画 */
    START_BY_CLICK(4, "START_BY_CLICK"),

    /** 从上一项开始触发动画 */
    START_BY_WITH(5, "START_BY_WITH"),

    /** 从上一项之后开始触发动画 */
    START_BY_AFTER(6, "START_BY_AFTER"),

    /** 动画延迟时间（ms） */
    DELAY(7, "DELAY"),

    /** 动画持续时间（ms) */
    DUR(8, "DUR");

    private final Integer id;
    private final String property;
}
