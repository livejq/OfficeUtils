package com.gzmut.office.enums;

import lombok.AccessLevel;
import lombok.AllArgsConstructor;
import lombok.Getter;

/**
 * 检查状态枚举
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-25
 */
@Getter
@AllArgsConstructor(access = AccessLevel.PRIVATE)
public enum CheckStatus implements CommonEnum {
    /**
     * 答案正确
     */
    CORRECT(0, "CORRECT", "正确"),

    /**
     * 答案错误
     */
    WRONG(1, "WRONG", "错误"),


    /**
     * 发生异常
     */
    EXCEPTION(2, "EXCEPTION", "异常");

    /**
     * 字段ID
     */
    private final Integer code;

    /**
     * 目标字段
     */
    private final String Name;

    /**
     * 字段描述
     */
    private final String desc;

}