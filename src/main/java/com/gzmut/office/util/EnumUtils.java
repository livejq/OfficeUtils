package com.gzmut.office.util;

import com.gzmut.office.enums.CommonEnum;


/**
 * 枚举工具类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-07-16
 */
public class EnumUtils {


    /**
     * 根据类和code值返回指定枚举对象
     *
     * @param clazz 所需查找的枚举类
     * @param code  code编码值
     * @return T
     */
    public static <T extends CommonEnum> T getEnumByCode(Class<T> clazz, Integer code) {
        for (T t : clazz.getEnumConstants()) {
            if (code.equals(t.getCode())) {
                return t;
            }
        }
        return null;
    }

    /**
     * 根据类和name值返回指定枚举对象
     *
     * @param clazz 所需查找的枚举类
     * @param name  枚举字段名
     * @return T
     */
    public static <T extends CommonEnum> T getEnumByName(Class<T> clazz, String name) {
        for (T t : clazz.getEnumConstants()) {
            if (name.equals(t.getName())) {
                return t;
            }
        }
        return null;
    }

    /**
     * 根据类和描述返回指定枚举对象
     *
     * @param clazz 所需查找的枚举类
     * @param desc  枚举描述
     * @return T
     */
    public static <T extends CommonEnum> T getEnumByDesc(Class<T> clazz, String desc) {
        for (T t : clazz.getEnumConstants()) {
            if (desc.equals(t.getDesc())) {
                return t;
            }
        }
        return null;
    }


}