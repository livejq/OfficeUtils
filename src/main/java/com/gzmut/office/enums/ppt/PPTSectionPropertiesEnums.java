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
public enum PPTSectionPropertiesEnums {

    /** 节名称 */
    SECTION_NAME(1, "SECTION_NAME"),

    /** 节对应的幻灯片数量 */
    SECTION_NUMBER(2, "SECTION_NUMBER");

    private final Integer id;
    private final String property;
}
