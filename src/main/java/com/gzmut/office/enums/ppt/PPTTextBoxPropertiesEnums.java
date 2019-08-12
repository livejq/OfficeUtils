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

    /** 粗体 */
    BOLD(6, "BOLD"),

    /** 斜体 */
    ITALIC(7, "ITALIC"),

    /** 下划线 */
    UNDERLINED(8, "UNDERLINED"),

    /** 脚注 */
    SUBSCRIPT(9, "SUBSCRIPT"),

    /** 段前间距 */
    SPACE_BEFORE(10, "SPACE_BEFORE"),

    /** 段后间距 */
    SPACE_AFTER(11, "SPACE_AFTER"),

    /** 行距 */
    LINE_SPACE(12, "LINE_SPACE"),

    /** 行高 */
    LINE_HEIGHT(13, "LINE_HEIGHT"),

    /** 标题 */
    TITLE(14, "TITLE"),

    /** 子标题 */
    SUBTITLE(15, "SUBTITLE"),

    /** 一级文本 */
    FIRST_TEXT(16, "FIRST_TEXT"),

    /** 二级文本 */
    SECOND_TEXT(17, "SECOND_TEXT"),

    /** 三级文本 */
    THIRD_TEXT(18, "THIRD_TEXT"),

    /** 文本超链接 */
    TEXT_HYPERLINK(19, "TEXT_HYPERLINK"),

    /** 文件超链接 */
    FILE_HYPERLINK(20, "FILE_HYPERLINK"),

    /** 网页超链接 */
    WEB_HYPERLINK(21, "WEB_HYPERLINK"),

    /** 水平对齐 相对于段落 algn="l" algn="ctr" algn="r"*/
    HORIZONTAL_ALIGN(22, "HORIZONTAL_ALIGN"),

    /** 垂直对齐 相对于文本框  anchor="t" anchor="ctr" anchor="b"*/
    VERTICAL_ALIGN(23, "VERTICAL_ALIGN");

    private final Integer id;
    private final String property;
}
