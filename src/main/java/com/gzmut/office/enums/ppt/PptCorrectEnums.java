package com.gzmut.office.enums.ppt;

import com.gzmut.office.enums.CommonEnum;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * @author livejq
 * @since 2019/7/27
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum PptCorrectEnums implements CommonEnum {

    /**
     * 检查幻灯片主题
     */
    CHECK_SLIDE_THEME(1, "CHECK_SLIDE_THEME", "检查幻灯片主题"),

    /**
     * 检查幻灯片数量
     */
    CHECK_SLIDE_NUMBER(2, "CHECK_SLIDE_NUMBER", "检查幻灯片数量"),

    /**
     * 检查幻灯片背景颜色
     */
    CHECK_SLIDE_BACKGROUND_COLOR(3, "CHECK_SLIDE_BACKGROUND_COLOR", "检查幻灯片背景颜色"),

    /**
     * 检查文本框内容
     */
    CHECK_TEXT_BOX_CONTENT(4, "CHECK_TEXT_BOX_CONTENT", "检查文本框内容"),

    /**
     * 检查文本框内容格式
     */
    CHECK_TEXT_BOX_CONTENT_FORMAT(5, "CHECK_TEXT_BOX_CONTENT_FORMAT", "检查文本框内容格式"),

    /**
     * 检查图片数量
     */
    CHECK_PICTURE_COUNT(6, "CHECK_PICTURE_COUNT", "检查图片数量"),

    /**
     * 检查版式
     */
    CHECK_LAYOUT_STYLE(7, "CHECK_LAYOUT_STYLE", "检查版式"),

    /**
     * 检查文本级别
     */
    CHECK_TEXT_LEVEL(8, "CHECK_TEXT_LEVEL", "检查文本级别"),

    /**
     * 检查母版名称
     */
    CHECK_LAYOUT_NAME(9, "CHECK_LAYOUT_NAME", "检查母版名称"),

    /**
     * 检查文本超链接
     */
    CHECK_TEXT_HYPERLINK(10, "CHECK_TEXT_HYPERLINK", "检查文本超链接"),

    /**
     * 检查文件超链接
     */
    CHECK_FILE_HYPERLINK(11, "CHECK_FILE_HYPERLINK", "检查文件超链接"),

    /**
     * 检查网页超链接
     */
    CHECK_WEB_HYPERLINK(12, "CHECK_WEB_HYPERLINK", "检查网页超链接"),

    /**
     * 检查文本对齐方式
     */
    CHECK_TEXT_ALIGN_STYLE(13, "CHECK_TEXT_ALIGN_STYLE", "检查文本对齐方式"),

    /**
     * 检查节
     */
    CHECK_SECTION(14, "CHECK_SECTION", "检查节"),

    /**
     * 检查幻灯片切换方式
     */
    CHECK_SWITCH_STYLE(15, "CHECK_SWITCH_STYLE", "检查幻灯片切换方式"),

    /**
     * 检查动画类型（entr/emph/exit）
     */
    CHECK_ANIMATION_TYPE(16, "CHECK_ANIMATION_TYPE", "检查动画类型"),

    /**
     * 检查文本框背景颜色
     */
    CHECK_TEXT_BOX_BACKGROUND_COLOR(17, "CHECK_TEXT_BOX_BACKGROUND_COLOR", "检查文本框背景颜色"),

    /**
     * 检查触发动画方式
     */
    CHECK_ANIMATION_START_STYLE(18, "CHECK_ANIMATION_START_STYLE", "检查触发动画方式"),

    /**
     * 检查动画延迟时间
     */
    CHECK_ANIMATION_DELAY(19, "CHECK_ANIMATION_DELAY", "检查动画延迟时间"),

    /**
     * 检查动画持续时间
     */
    CHECK_ANIMATION_DUR(20, "CHECK_ANIMATION_DUR", "检查动画持续时间"),

    /**
     * 检查版式
     */
    FORMAT(21, "FORMAT", "检查版式"),

    /**
     * 检查文本内容
     */
    TEXT_CONTENT(22, "TEXT_CONTENT", "检查文本内容"),

    /**
     * 检查样式
     */
    STYLE(23, "STYLE", "检查样式"),

    /**
     * 检查颜色
     */
    COLOR(24, "COLOR", "检查颜色"),

    /**
     * 检查字体
     */
    FONT(25, "FONT", "检查字体");

    /**
     * 属性id
     */
    private final Integer code;

    /**
     * 检查方法名称
     */
    private final String name;

    /**
     * 检查描述
     */
    private final String desc;
}
