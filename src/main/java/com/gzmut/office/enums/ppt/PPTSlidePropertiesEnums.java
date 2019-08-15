package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.logging.Level;

/**
 * 封装幻灯片属性
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPTSlidePropertiesEnums {

    /** 背景颜色 */
    BACKGROUND_COLOR(1, "BACKGROUND_COLOR"),

    /** 切入方式 */
    SWITCH_STYLE(2, "SWITCH_STYLE"),

    /** 主题 */
    THEME(3, "THEME"),

    /** 板式类型 */
    LAYOUT_STYLE(4, "LAYOUT_STYLE"),

    /** 母版名称 */
    LAYOUT_NAME(5, "LAYOUT_NAME"),

    /** 尺寸 */
    SIZE(6, "SIZE"),

    /** 标题 */
    TITLE(7, "TITLE"),

    /** 幻灯片数量 */
    COUNT(8, "COUNT");

    private final Integer id;
    private final String property;
}
