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
public enum PPtPicturePropertiesEnums {

    /** 尺寸 */
    SIZE(1, "SIZE"),

    /** 超链接 */
    HYPERLINK(2, "HYPERLINK"),

    /** 类型 */
    CONTENT_TYPE(3, "CONTENT_TYPE"),

    /** 后缀 */
    EXTENSION(4, "EXTENSION");

    private final Integer id;
    private final String property;
}
