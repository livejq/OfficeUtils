package com.gzmut.office.enums.ppt;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author livejq
 * @since 2019/7/27
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PPTCorrectEnums {

    /** 检查幻灯片主题 */
    CHECK_SLIDE_THEME(1,"CHECK_SLIDE_THEME", "检查幻灯片主题"),

    /** 检查目标数量 */
    CHECK_SLIDE_NUMBER(2, "CHECK_SLIDE_NUMBER", "检查幻灯片数量"),

    /** 检查背景颜色 */
    CHECK_BACKGROUND_COLOR(3, "CHECK_BACKGROUND_COLOR", "检查背景颜色"),

    /** 检查文本框内容 */
    CHECK_TEXT_BOX_CONTENT(4, "CHECK_TEXT_BOX_CONTENT", "检查文本框内容"),

    /** 检查段落样式 */
    CHECK_TEXT_BOX_CONTENT_FORMAT(5, "CHECK_TEXT_BOX_CONTENT_FORMAT", "检查文本框内容格式"),

    /** 检查段落样式 */
    CHECK_PARAGRAPH_STYLE(6,"CHECK_PARAGRAPH_STYLE", "检查段落样式"),

    /** 检查版式 */
    CHECK_LAYOUT_STYLE(7,"CHECK_LAYOUT_STYLE", "检查版式"),

    /** 检查文本级别 */
    CHECK_TEXT_LEVEL(8,"CHECK_TEXT_LEVEL", "检查文本级别"),

    /** 检查母版名称 */
    CHECK_LAYOUT_NAME(9,"CHECK_LAYOUT_NAME", "检查母版名称"),

    /** 检查文本超链接 */
    CHECK_TEXT_HYPERLINK(10,"CHECK_TEXT_HYPERLINK", "检查文本超链接"),

    /** 检查文件超链接 */
    CHECK_FILE_HYPERLINK(11,"CHECK_FILE_HYPERLINK", "检查文件超链接"),

    /** 检查网页超链接 */
    CHECK_WEB_HYPERLINK(12,"CHECK_WEB_HYPERLINK", "检查网页超链接"),

    /** 检查文本对齐方式 */
    CHECK_TEXT_ALIGN_STYLE(13,"CHECK_TEXT_ALIGN_STYLE", "检查文本对齐方式"),

    /** 检查节 */
    CHECK_SECTION(14,"CHECK_SECTION", "检查节");

    /** 属性id */
    private final Integer id;
    /** 检查方法名称 */
    private final String checkName;
    /** 检索信息 */
    private final String msg;
}
