package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 封装文本框属性
 * @author livejq
 * @date 2019/7/13
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPtTextBoxPropertiesEnums {

    /** 填充颜色 */
    FILL_COLOR(1, "填充颜色"),

    /** 文本内容 */
    TEXT_CONTENT(2, "文本内容"),

    /** 字体样式 */
    TEXT_STYLE(3, "字体样式"),

    /** 字体大小 */
    TEXT_SIZE(4, "字体大小"),

    /** 字体颜色 */
    TEXT_COLOR(5, "字体颜色"),

    /** 超链接 */
    HYPERLINK(6, "超链接"),

    /** 粗体 */
    BOLD(7, "粗体"),

    /** 斜体 */
    ITALIC(8, "斜体"),

    /** 下划线 */
    UNDERLINED(9, "下划线"),

    /** 脚注 */
    SUBSCRIPT(10, "脚注"),

    /** 段前间距 */
    SPACE_BEFORE(11, "段前间距"),

    /** 段后间距 */
    SPACE_AFTER(12, "段后间距"),

    /** 行距 */
    LINE_SPACE(13, "行距"),

    /** 行高 */
    LINE_HEIGHT(14, "行高");

    private final Integer id;
    private final String property;
}
