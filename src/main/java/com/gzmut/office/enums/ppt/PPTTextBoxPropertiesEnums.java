package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 封装文本框属性
 * @author livejq
 * @since 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPTTextBoxPropertiesEnums {

    /** 填充颜色 */
    FILL_COLOR(1, "FILL_COLOR"),

    /** 文本内容 */
    TEXT_CONTENT(2, "TEXT_CONTENT"),

    /** 字体样式 */
    TEXT_STYLE(3, "TEXT_STYLE"),

    /** 字体大小 */
    TEXT_SIZE(4, "TEXT_SIZE"),

    /** 字体颜色 */
    TEXT_COLOR(5, "TEXT_COLOR"),

    /** 超链接 */
    HYPERLINK(6, "HYPERLINK"),

    /** 粗体 */
    BOLD(7, "BOLD"),

    /** 斜体 */
    ITALIC(8, "ITALIC"),

    /** 下划线 */
    UNDERLINED(9, "UNDERLINED"),

    /** 脚注 */
    SUBSCRIPT(10, "SUBSCRIPT"),

    /** 段前间距 */
    SPACE_BEFORE(11, "SPACE_BEFORE"),

    /** 段后间距 */
    SPACE_AFTER(12, "SPACE_AFTER"),

    /** 行距 */
    LINE_SPACE(13, "LINE_SPACE"),

    /** 行高 */
    LINE_HEIGHT(14, "LINE_HEIGHT");

    private final Integer id;
    private final String property;
}
