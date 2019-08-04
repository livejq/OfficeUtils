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
public enum PPtCorrectEnums {

    /** 检查文件是否存在 */
    CHECK_FILE_IS_EXISTS(1, "CHECK_FILE_IS_EXISTS", "检查文件是否存在"),

    /** 检查目标数量 */
    CHECK_TARGET_NUMBER(2, "CHECK_TARGET_NUMBER", "检查目标数量"),

    /** 检查背景颜色 */
    CHECK_BACKGROUND_COLOR(3, "CHECK_BACKGROUND_COLOR", "检查背景颜色"),

    /** 检查文本框内容 */
    CHECK_TEXT_BOX_CONTENT(4, "CHECK_TEXT_BOX_CONTENT", "检查文本框内容"),

    /** 检查段落样式 */
    CHECK_TEXT_BOX_CONTENT_FORMAT(5, "CHECK_TEXT_BOX_CONTENT_FORMAT", "检查文本框内容格式"),

    /** 检查段落样式 */
    CHECK_PARAGRAPH_STYLE(6,"CHECK_PARAGRAPH_STYLE", "检查段落样式"),

    /** 检查版式 */
    CHECK_LAYOUT_STYLE(6,"CHECK_LAYOUT_STYLE", "检查版式");

    /** 属性id */
    private final Integer id;
    /** 检查方法名称 */
    private final String checkName;
    /** 检索信息 */
    private final String msg;
}
