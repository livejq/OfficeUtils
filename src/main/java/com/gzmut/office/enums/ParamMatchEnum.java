package com.gzmut.office.enums;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 参数匹配模式枚举
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-25
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum ParamMatchEnum implements CommonEnum {
    /**
     * 匹配参数模式
     */
    INCLUDE(0, "INCLUDE", "匹配参数模式"),

    /**
     * 排除参数模式
     */
    EXCLUDE(1, "EXCLUDE", "排除参数模式");

    /**
     * 参数匹配ID
     */
    private final Integer code;

    /**
     * 参数匹配字段
     */
    private final String Name;

    /**
     * 字段描述
     */
    private final String desc;
}