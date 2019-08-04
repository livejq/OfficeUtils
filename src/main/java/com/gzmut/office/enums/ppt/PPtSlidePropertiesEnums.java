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
public enum PPtSlidePropertiesEnums {

    /** 背景颜色 */
    BACKGROUND_COLOR(1, "BACKGROUND_COLOR"),

    /** 切入方式 */
    SLIP_STYLE(2, "SLIP_STYLE"),

    /** 主题 */
    THEME(3, "THEME"),

    /** 布局类型 */
    LAYOUT_STYLE(4, "LAYOUT_STYLE"),

    /** 尺寸 */
    SIZE(5, "SIZE"),

    /** 标题 */
    TITLE(6, "TITLE"),

    /** 标题 */
    COUNT(7, "COUNT");

    private final Integer id;
    private final String property;
}
