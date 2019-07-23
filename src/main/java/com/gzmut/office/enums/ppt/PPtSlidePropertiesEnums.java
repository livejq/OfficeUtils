package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.Setter;

import java.util.logging.Level;

/**
 * 封装幻灯片属性
 * @author livejq
 * @date 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPtSlidePropertiesEnums {

    /** 背景颜色 */
    BACKGROUND_COLOR(1, "背景颜色"),

    /** 切入方式 */
    SLIP_STYLE(2, "切入方式"),

    /** 主题 */
    THEME(3, "主题"),

    /** 布局类型 */
    LAYOUT_STYLE(4, "布局类型"),

    /** 尺寸 */
    SIZE(5, "尺寸"),

    /** 标题 */
    TITLE(6, "标题");

    private final Integer id;
    private final String property;
}
