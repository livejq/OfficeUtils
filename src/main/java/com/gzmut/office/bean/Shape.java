package com.gzmut.office.bean;

import lombok.Data;
import lombok.experimental.Accessors;

/**
 * 形状实体类
 *
 * @author zzzzzzzzzzzzzzz
 * @date 2019-08-08
 */
@Data
@Accessors(chain = true)
public class Shape {
    /** 形状元素自带ID */
    private String id;

    /** 形状元素自带name */
    private String name;

    /** 形状元素内文本内容 */
    private String text;

    /** 形状元素轮廓主题 */
    private String lnVal;

    /** 形状元素填充主题 */
    private String fillVal;

    /** 形状效果主题 */
    private String effectVal;

    /** 形状字体主题 */
    private String fontVal;

    /**
     * 判断当前形状对象是否为空
     *
     * @return boolean
     */
    public boolean isNotBlank() {
        return (id != null && name != null);
    }
}